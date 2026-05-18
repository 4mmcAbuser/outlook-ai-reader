// Initialize Lucide Icons
lucide.createIcons();

// --- STATE MANAGEMENT ---
let config = {
    apiKey: '',
    model: 'gemini-2.5-flash-lite',
    tone: 'Επαγγελματικός, ευγενικός και σοβαρός.',
    autoSum: true,
    customPrompt: ''
};
let emailContext = { text: '', meta: {}, fullConversation: [] };
let isRecording = false;
let mediaRecorder;
let audioChunks = [];

// --- INITIALIZATION ---
Office.onReady((info) => {
    loadSettings();
    if (info.host === Office.HostType.Outlook) {
        initOutlookData();
        // Παρακολούθηση αλλαγής επιλεγμένου email
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
            console.log("Item changed, reloading conversation...");
            setTimeout(() => initOutlookData(), 200);
        });
    }
});

function loadSettings() {
    const saved = localStorage.getItem('aiAssistConfig');
    if (saved) config = { ...config, ...JSON.parse(saved) };
    
    document.getElementById('setApiKey').value = config.apiKey;
    document.getElementById('setModel').value = config.model;
    document.getElementById('setTone').value = config.tone;
    document.getElementById('setAutoSum').checked = config.autoSum;
    document.getElementById('setCustomPrompt').value = config.customPrompt;

    if (!config.apiKey) navigate('view-settings');
}

function saveSettings() {
    config.apiKey = document.getElementById('setApiKey').value.trim();
    config.model = document.getElementById('setModel').value;
    config.tone = document.getElementById('setTone').value;
    config.autoSum = document.getElementById('setAutoSum').checked;
    config.customPrompt = document.getElementById('setCustomPrompt').value.trim();
    
    localStorage.setItem('aiAssistConfig', JSON.stringify(config));
    navigate('view-main');
    
    if (config.autoSum && !emailContext.text) initOutlookData();
}

function navigate(viewId) {
    document.querySelectorAll('.view-section').forEach(el => {
        el.classList.remove('view-active');
    });
    document.getElementById(viewId).classList.add('view-active');
}

// --- EWS: ΑΝΑΚΤΗΣΗ ΟΛΟΚΛΗΡΗΣ ΣΥΝΟΜΙΛΙΑΣ (χωρίς backend) ---
async function getFullConversationViaEWS() {
    return new Promise((resolve, reject) => {
        const itemId = Office.context.mailbox.item.itemId;
        const conversationId = Office.context.mailbox.item.conversationId;
        const ewsUrl = Office.context.mailbox.ewsUrl;
        
        if (!ewsUrl || !conversationId) {
            reject(new Error("Missing EWS URL or Conversation ID"));
            return;
        }

        // 1. FindItem SOAP request για να βρούμε όλα τα items της συνομιλίας
        const findItemSoap = `<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
            <soap:Header>
                <t:RequestServerVersion Version="Exchange2013" />
            </soap:Header>
            <soap:Body>
                <m:FindItem Traversal="Shallow">
                    <m:ItemShape>
                        <t:BaseShape>IdOnly</t:BaseShape>
                        <t:AdditionalProperties>
                            <t:FieldURI FieldURI="item:Subject" />
                            <t:FieldURI FieldURI="item:DateTimeReceived" />
                            <t:FieldURI FieldURI="item:Body" />
                            <t:FieldURI FieldURI="item:Sender" />
                        </t:AdditionalProperties>
                    </m:ItemShape>
                    <m:IndexedPageItemView MaxEntriesReturned="100" Offset="0" BasePoint="Beginning" />
                    <m:ParentFolderIds>
                        <t:DistinguishedFolderId Id="inbox" />
                    </m:ParentFolderIds>
                    <m:Restriction>
                        <t:IsEqualTo>
                            <t:FieldURI FieldURI="item:ConversationId" />
                            <t:FieldURIOrConstant>
                                <t:Constant Value="${conversationId}" />
                            </t:FieldURIOrConstant>
                        </t:IsEqualTo>
                    </m:Restriction>
                </m:FindItem>
            </soap:Body>
        </soap:Envelope>`;

        // Κάνουμε το SOAP request
        $.ajax({
            url: ewsUrl,
            type: 'POST',
            data: findItemSoap,
            contentType: 'text/xml; charset=utf-8',
            dataType: 'xml',
            success: function(xmlDoc) {
                const items = [];
                $(xmlDoc).find('Items > Item').each(function() {
                    const itemIdVal = $(this).find('ItemId').attr('Id');
                    const changeKey = $(this).find('ItemId').attr('ChangeKey');
                    const subject = $(this).find('Subject').text();
                    const dateTime = $(this).find('DateTimeReceived').text();
                    const senderName = $(this).find('Sender Mailbox Name').text();
                    const senderEmail = $(this).find('Sender Mailbox EmailAddress').text();
                    // Χρειαζόμαστε και το Body - θα το πάρουμε ξεχωριστά με GetItem
                    items.push({
                        itemId: itemIdVal,
                        changeKey: changeKey,
                        subject: subject,
                        dateTime: dateTime,
                        senderName: senderName,
                        senderEmail: senderEmail
                    });
                });

                // 2. Για κάθε item, κάνουμε GetItem για να πάρουμε το πλήρες σώμα
                const promises = items.map(item => {
                    return new Promise((resolveItem) => {
                        const getItemSoap = `<?xml version="1.0" encoding="utf-8"?>
                        <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
                            <soap:Header>
                                <t:RequestServerVersion Version="Exchange2013" />
                            </soap:Header>
                            <soap:Body>
                                <m:GetItem>
                                    <m:ItemShape>
                                        <t:BaseShape>AllProperties</t:BaseShape>
                                    </m:ItemShape>
                                    <m:ItemIds>
                                        <t:ItemId Id="${item.itemId}" ChangeKey="${item.changeKey}" />
                                    </m:ItemIds>
                                </m:GetItem>
                            </soap:Body>
                        </soap:Envelope>`;

                        $.ajax({
                            url: ewsUrl,
                            type: 'POST',
                            data: getItemSoap,
                            contentType: 'text/xml; charset=utf-8',
                            dataType: 'xml',
                            success: function(getXml) {
                                let bodyText = '';
                                const bodyElem = $(getXml).find('Body');
                                if (bodyElem.length) {
                                    let rawBody = bodyElem.text();
                                    // Καθαρισμός HTML → plain text
                                    let cleanBody = rawBody
                                        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                                        .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                                        .replace(/<br\s*[\/]?>/gi, '\n')
                                        .replace(/<\/p>/gi, '\n\n')
                                        .replace(/<div[^>]*>/gi, '')
                                        .replace(/<\/div>/gi, '\n')
                                        .replace(/<[^>]+>/g, '');
                                    const txtDiv = document.createElement('textarea');
                                    txtDiv.innerHTML = cleanBody;
                                    bodyText = txtDiv.value.trim();
                                }
                                resolveItem({
                                    ...item,
                                    body: bodyText
                                });
                            },
                            error: () => resolveItem({ ...item, body: '' })
                        });
                    });
                });

                Promise.all(promises).then(fullItems => {
                    // Ταξινόμηση κατά ημερομηνία (παλιότερο πρώτο)
                    fullItems.sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));
                    resolve(fullItems);
                }).catch(reject);
            },
            error: function(err) {
                reject(err);
            }
        });
    });
}

// --- ΚΥΡΙΑ ΣΥΝΑΡΤΗΣΗ ΦΟΡΤΩΣΗΣ EMAIL ---
function initOutlookData() {
    const item = Office.context.mailbox.item;
    
    // Metadata (ίδια όπως πριν)
    emailContext.meta = {
        senderName: item.sender ? item.sender.displayName : "Άγνωστος",
        senderEmail: item.sender ? item.sender.emailAddress : "unknown@mail.com",
        subject: item.subject || "(Χωρίς θέμα)",
        toRecipients: item.to ? item.to.map(r => r.displayName).join(", ") : "Μόνο εγώ",
        receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('el-GR') : "Άγνωστη ημερομηνία"
    };

    // Πρώτα δοκιμάζουμε να πάρουμε ολόκληρη τη συνομιλία μέσω EWS
    getFullConversationViaEWS()
        .then(conversationItems => {
            console.log(`Λήφθηκαν ${conversationItems.length} μηνύματα από EWS`);
            emailContext.fullConversation = conversationItems;
            
            // Συγχώνευση όλων των σωμάτων σε ένα κείμενο
            let fullText = conversationItems.map(msg => {
                return `--- Μήνυμα από: ${msg.senderName || msg.senderEmail} (${new Date(msg.dateTime).toLocaleString('el-GR')}) ---\nΘέμα: ${msg.subject}\n\n${msg.body}\n\n`;
            }).join('\n');
            
            emailContext.text = fullText;
            
            // Αν το EWS απέτυχε ή δεν επέστρεψε τίποτα, κάνουμε fallback στο συνηθισμένο body
            if (!emailContext.text || emailContext.text.length < 50) {
                console.log("EWS didn't return enough content, falling back to simple body");
                getSimpleBody();
            } else {
                finishLoading();
            }
        })
        .catch(err => {
            console.warn("EWS error, falling back to simple body", err);
            getSimpleBody();
        });
    
    function getSimpleBody() {
        if (item.body) {
            item.body.getAsync(Office.CoercionType.Html, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    let rawHtml = result.value;
                    // Αφαίρεση display:none και άλλων περιορισμών
                    rawHtml = rawHtml.replace(/display\s*:\s*none/g, 'display:block');
                    rawHtml = rawHtml.replace(/visibility\s*:\s*hidden/g, 'visibility:visible');
                    let cleanText = rawHtml
                        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                        .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                        .replace(/<br\s*[\/]?>/gi, '\n')
                        .replace(/<\/p>/gi, '\n\n')
                        .replace(/<\/div>/gi, '\n')
                        .replace(/<div[^>]*>/gi, '')
                        .replace(/<[^>]+>/g, '');
                    const txt = document.createElement('textarea');
                    txt.innerHTML = cleanText;
                    emailContext.text = txt.value.trim();
                } else {
                    item.body.getAsync(Office.CoercionType.Text, (textResult) => {
                        if (textResult.status === Office.AsyncResultStatus.Succeeded) {
                            emailContext.text = textResult.value;
                        }
                    });
                }
                finishLoading();
            });
        } else {
            finishLoading();
        }
    }
}

function finishLoading() {
    if (config.autoSum && config.apiKey && emailContext.text && emailContext.text.length > 50) {
        generateSummary();
    } else if (!config.autoSum) {
        showManualSummaryBtn();
    } else {
        if (emailContext.text && emailContext.text.length < 50) {
            document.getElementById('summaryText').innerHTML = "⚠️ Δεν ήταν δυνατή η λήψη ολόκληρης της συνομιλίας.<br>Δοκιμάστε να ανοίξετε το email σε νέο παράθυρο.";
        }
        showManualSummaryBtn();
        stopLoadingAnim();
    }
}

// --- FAKE LOADING ANIMATION (ίδια) ---
let loadingInterval;
function startLoadingAnim(messages) {
    const textEl = document.getElementById('loadingText');
    document.getElementById('loadingOverlay').style.display = 'flex';
    document.getElementById('summaryText').style.opacity = '0.3';
    let i = 0;
    textEl.innerText = messages[0];
    loadingInterval = setInterval(() => {
        i = (i + 1) % messages.length;
        textEl.innerText = messages[i];
    }, 1500);
}
function stopLoadingAnim() {
    clearInterval(loadingInterval);
    document.getElementById('loadingOverlay').style.display = 'none';
    document.getElementById('summaryText').style.opacity = '1';
}

function showManualSummaryBtn() {
    document.getElementById('loadingOverlay').style.display = 'none';
    document.getElementById('summaryText').innerText = "Η αυτόματη σύνοψη είναι ανενεργή ή δεν βρέθηκε περιεχόμενο.";
    const btn = document.getElementById('manualSummaryBtn');
    btn.style.display = 'block';
    btn.onclick = () => { btn.style.display = 'none'; generateSummary(); };
}

// --- SUMMARY GENERATION (ίδια) ---
async function generateSummary() {
    if (!config.apiKey) return;
    startLoadingAnim(["Διαβάζω το Thread...", "Αναλύω τη συνομιλία...", "Ετοιμάζω σύνοψη..."]);
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
    const prompt = `Κάνε μια πολύ σύντομη, περιεκτική σύνοψη (max 3-4 γραμμές) στα Ελληνικά για την παρακάτω συνομιλία email. Ποιος στέλνει το τελευταίο μήνυμα και τι ζητάει.\nΣυνομιλία:\n${emailContext.text.substring(0, 15000)}`; // περιορισμός length
    
    try {
        const res = await fetch(url, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
        });
        const data = await res.json();
        stopLoadingAnim();
        if (data.error) throw new Error(data.error.message);

        const summary = data.candidates[0].content.parts[0].text;
        const sumEl = document.getElementById('summaryText');
        sumEl.innerText = summary;
        
        if (sumEl.scrollHeight > 96) { 
            document.getElementById('summaryFade').style.display = 'block';
            const expBtn = document.getElementById('expandSummaryBtn');
            expBtn.style.display = 'block';
            expBtn.onclick = () => {
                sumEl.parentElement.classList.remove('max-h-24');
                document.getElementById('summaryFade').style.display = 'none';
                expBtn.style.display = 'none';
            };
        }
    } catch (e) {
        stopLoadingAnim();
        document.getElementById('summaryText').innerText = "Σφάλμα κατά τη σύνοψη: " + e.message;
    }
}

// --- ACTION LOGIC (Text, Quick Actions, Voice) ---
function handleQuickAction(actionType) {
    if(!config.apiKey) return alert("Βάλε API Key στις ρυθμίσεις");
    generateDraft(actionType, null);
}

document.getElementById('sendTextBtn').onclick = () => {
    const input = document.getElementById('textPrompt').value.trim();
    if(input && config.apiKey) {
        document.getElementById('textPrompt').value = '';
        generateDraft(input, null);
    }
};

document.getElementById('tweakBtn').onclick = () => {
    const tweak = document.getElementById('tweakPrompt').value.trim();
    if(tweak) {
        const currentDraft = document.getElementById('draftTextarea').value;
        document.getElementById('tweakPrompt').value = '';
        generateDraft(`ΟΔΗΓΙΑ ΤΡΟΠΟΠΟΙΗΣΗΣ EMAIL: "${tweak}". Τροποποίησε αυτό το κείμενο που έγραψες: "${currentDraft}"`, null);
    }
}

const voiceBtn = document.getElementById('voiceBtn');
const voiceStatus = document.getElementById('voiceStatus');

voiceBtn.onclick = () => {
    if(!config.apiKey) return alert("Βάλε API Key στις ρυθμίσεις");
    if (isRecording) { stopRecording(); return; }
    
    if (Office.context.mailbox && Office.devicePermission) {
        Office.devicePermission.requestPermissionsAsync([Office.DevicePermissionType.microphone], (res) => {
            if (res.value) location.reload(); 
            else if (res.status !== Office.AsyncResultStatus.Failed) startRecording();
        });
    } else {
        startRecording();
    }
};

function startRecording() {
    navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        mediaRecorder = new MediaRecorder(stream);
        audioChunks = [];
        
        mediaRecorder.ondataavailable = e => { if (e.data.size > 0) audioChunks.push(e.data); };
        mediaRecorder.onstop = () => {
            voiceStatus.innerText = "Επεξεργασία ήχου...";
            const blob = new Blob(audioChunks, { type: mediaRecorder.mimeType || 'audio/webm' });
            const reader = new FileReader();
            reader.readAsDataURL(blob);
            reader.onloadend = () => {
                const b64 = reader.result.split(',')[1];
                const mime = (mediaRecorder.mimeType || 'audio/webm').split(';')[0];
                generateDraft("Αυτό είναι ηχητικό μήνυμα.", { data: b64, mimeType: mime });
            };
        };

        mediaRecorder.start();
        isRecording = true;
        voiceBtn.classList.remove('siri-idle');
        voiceBtn.classList.add('siri-listening');
        voiceBtn.innerHTML = `<i data-lucide="square" class="w-8 h-8 text-white"></i>`;
        lucide.createIcons();
        voiceStatus.innerText = "Ακούω... Πατήστε για τέλος";
    }).catch(err => {
        voiceStatus.innerText = "Σφάλμα μικροφώνου!";
    });
}

function stopRecording() {
    if (mediaRecorder && mediaRecorder.state !== "inactive") mediaRecorder.stop();
    isRecording = false;
    voiceBtn.classList.remove('siri-listening');
    voiceBtn.classList.add('siri-idle');
    voiceBtn.innerHTML = `<i data-lucide="mic" class="w-8 h-8 opacity-70"></i>`;
    lucide.createIcons();
}

// --- CORE GENERATOR ENGINE (INTENT AWARE) ---
async function generateDraft(instruction, audioObj) {
    voiceStatus.innerText = "Σκέφτομαι...";
    
    const systemPrompt = `Είσαι ένας έμπειρος Executive Assistant. Έχεις μπροστά σου το ιστορικό μιας συνομιλίας (Email Thread). 
Τα πιο πρόσφατα μηνύματα είναι στην κορυφή, τα παλαιότερα στο κάτω μέρος.

ΣΤΟΙΧΕΙΑ ΜΗΝΥΜΑΤΟΣ: 
Αποστολέας: ${emailContext.meta.senderName} (${emailContext.meta.senderEmail})
Θέμα: ${emailContext.meta.subject}

ΙΣΤΟΡΙΚΟ ΣΥΝΟΜΙΛΙΑΣ:
"""
${emailContext.text.substring(0, 12000)}
"""

ΟΔΗΓΙΑ/ΕΡΩΤΗΣΗ ΧΡΗΣΤΗ (Από ήχο ή κείμενο): ${instruction}

ΚΑΘΗΚΟΝ:
Πρέπει να καταλάβεις αν ο χρήστης σου ζητάει να ΓΡΑΨΕΙΣ ΜΙΑ ΑΠΑΝΤΗΣΗ (π.χ. "δέξου το", "απάντα ότι", "γράψε ένα mail") ή αν σου κάνει μια ΕΡΩΤΗΣΗ/ΣΥΖΗΤΗΣΗ (π.χ. "τι λέει εδώ;", "ποιος είναι;", "τι έγινε;").

Πρέπει ΑΥΣΤΗΡΑ να απαντήσεις ΜΟΝΟ με ένα JSON αντικείμενο, ακριβώς σε αυτή τη μορφή:
{
  "intent": "draft",
  "content": "Το κείμενο σου εδώ"
}
ή 
{
  "intent": "question",
  "content": "Η απάντησή σου εδώ"
}

ΚΑΝΟΝΕΣ ΑΝ intent == "draft":
- Γράψε ΑΠΕΥΘΕΙΑΣ το κείμενο του email προς τον πελάτη.
- ΤΟΝΟΣ: ${config.tone}. EXTRA ΚΑΝΟΝΕΣ: ${config.customPrompt}.

ΚΑΝΟΝΕΣ ΑΝ intent == "question":
- Δώσε μια ξεκάθαρη, φιλική απάντηση στον χρήστη. 
- ΜΗΝ το γράφεις σαν email.`;

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
    
    let parts = [{ text: systemPrompt }];
    if (audioObj) {
        parts.push({ inlineData: { mimeType: audioObj.mimeType, data: audioObj.data } });
    }

    try {
        const res = await fetch(url, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ 
                contents: [{ parts: parts }],
                generationConfig: { responseMimeType: "application/json" } 
            })
        });
        
        const data = await res.json();
        if(data.error) throw new Error(data.error.message);
        
        const rawResponse = data.candidates[0].content.parts[0].text.trim();
        const cleanJsonString = rawResponse.replace(/```json/gi, '').replace(/```/g, '').trim();
        const parsedResponse = JSON.parse(cleanJsonString);
        
        if (parsedResponse.intent === "question") {
            document.getElementById('answerText').innerText = parsedResponse.content;
            navigate('view-answer');
        } else {
            document.getElementById('draftTextarea').value = parsedResponse.content;
            navigate('view-draft');
        }
        
        voiceStatus.innerText = "Κάντε κλικ για ομιλία";
        
    } catch (e) {
        console.error(e);
        alert("Σφάλμα AI: Βεβαιωθείτε ότι το Prompt ήταν ξεκάθαρο.");
        voiceStatus.innerText = "Σφάλμα. Προσπαθήστε ξανά.";
    }
}

// --- INSERT TO OUTLOOK ---
document.getElementById('insertOutlookBtn').onclick = () => {
    const finalTxt = document.getElementById('draftTextarea').value;
    Office.context.mailbox.item.displayReplyForm(finalTxt);
    document.getElementById('draftTextarea').value = '';
    navigate('view-main');
};

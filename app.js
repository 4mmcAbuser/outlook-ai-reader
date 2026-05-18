// Initialize Lucide Icons
lucide.createIcons();

// ----------------------
// STATE MANAGEMENT
// ----------------------
let config = {
    apiKey: '',
    model: 'gemini-2.5-flash-lite',
    tone: 'Επαγγελματικός, ευγενικός και σοβαρός.',
    autoSum: true,
    customPrompt: ''
};

let emailContext = {
    text: '',
    meta: {},
    fullConversation: []
};

let isRecording = false;
let mediaRecorder;
let audioChunks = [];

// ----------------------
// INITIALIZATION
// ----------------------
Office.onReady((info) => {
    loadSettings();

    if (info.host === Office.HostType.Outlook) {
        initOutlookData();

        // Ανανέωση όταν ο χρήστης αλλάζει email στο Outlook
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            () => {
                console.log("Item changed");
                setTimeout(() => initOutlookData(), 300);
            }
        );
    }
});

// ----------------------
// SETTINGS
// ----------------------
function loadSettings() {
    const saved = localStorage.getItem('aiAssistConfig');

    if (saved) {
        config = { ...config, ...JSON.parse(saved) };
    }

    document.getElementById('setApiKey').value = config.apiKey;
    document.getElementById('setModel').value = config.model;
    document.getElementById('setTone').value = config.tone;
    document.getElementById('setAutoSum').checked = config.autoSum;
    document.getElementById('setCustomPrompt').value = config.customPrompt;

    if (!config.apiKey) {
        navigate('view-settings');
    }
}

function saveSettings() {
    config.apiKey = document.getElementById('setApiKey').value.trim();
    config.model = document.getElementById('setModel').value;
    config.tone = document.getElementById('setTone').value;
    config.autoSum = document.getElementById('setAutoSum').checked;
    config.customPrompt = document.getElementById('setCustomPrompt').value.trim();

    localStorage.setItem('aiAssistConfig', JSON.stringify(config));

    navigate('view-main');

    if (config.autoSum) {
        initOutlookData();
    }
}

// ----------------------
// NAVIGATION
// ----------------------
function navigate(viewId) {
    document.querySelectorAll('.view-section').forEach(el => {
        el.classList.remove('view-active');
    });
    document.getElementById(viewId).classList.add('view-active');
}

// ----------------------
// CLEAN HTML (Για το EWS)
// ----------------------
function cleanHtmlToText(html) {
    if (!html) return '';
    return html
        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
        .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>/gi, '\n\n')
        .replace(/<\/div>/gi, '\n')
        .replace(/<div[^>]*>/gi, '')
        .replace(/<[^>]+>/g, '')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}

// ----------------------
// GET FULL THREAD VIA EWS (Plan A)
// ----------------------
async function getFullConversationViaEWS() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;

        if (!item) {
            reject(new Error("No item"));
            return;
        }

        item.getAllInternetHeadersAsync((headerResult) => {
            if (headerResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("Cannot read headers"));
                return;
            }

            const headers = headerResult.value || '';
            const threadIndexMatch = headers.match(/Thread-Index:\s*(.+)/i);

            if (!threadIndexMatch) {
                reject(new Error("No Thread-Index found"));
                return;
            }

            const threadIndex = threadIndexMatch[1].trim();

            const soap = `
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2016"/>
    </soap:Header>
    <soap:Body>
        <m:FindItem Traversal="Deep">
            <m:ItemShape>
                <t:BaseShape>AllProperties</t:BaseShape>
                <t:AdditionalProperties>
                    <t:FieldURI FieldURI="item:Subject"/>
                    <t:FieldURI FieldURI="item:DateTimeReceived"/>
                    <t:FieldURI FieldURI="item:Body"/>
                    <t:FieldURI FieldURI="message:From"/>
                    <t:FieldURI FieldURI="message:Sender"/>
                    <t:FieldURI FieldURI="item:ConversationId"/>
                </t:AdditionalProperties>
            </m:ItemShape>
            <m:IndexedPageItemView MaxEntriesReturned="100" Offset="0" BasePoint="Beginning"/>
            <m:Restriction>
                <t:Contains ContainmentMode="Substring" ContainmentComparison="IgnoreCase">
                    <t:FieldURI FieldURI="item:ConversationIndex"/>
                    <t:Constant Value="${threadIndex.substring(0, 22)}"/>
                </t:Contains>
            </m:Restriction>
            <m:ParentFolderIds>
                <t:DistinguishedFolderId Id="msgfolderroot"/>
            </m:ParentFolderIds>
        </m:FindItem>
    </soap:Body>
</soap:Envelope>`;

            Office.context.mailbox.makeEwsRequestAsync(soap, (result) => {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(result.error);
                    return;
                }

                const parser = new DOMParser();
                const xml = parser.parseFromString(result.value, "text/xml");
                const messageNodes = xml.getElementsByTagName("t:Message");
                const messages = [];

                for (let i = 0; i < messageNodes.length; i++) {
                    const msg = messageNodes[i];
                    const subject = msg.getElementsByTagName("t:Subject")[0]?.textContent || '';
                    const body = msg.getElementsByTagName("t:Body")[0]?.textContent || '';
                    const received = msg.getElementsByTagName("t:DateTimeReceived")[0]?.textContent || '';
                    let sender = 'Unknown';
                    const fromNode = msg.getElementsByTagName("t:From")[0];

                    if (fromNode) {
                        sender = fromNode.getElementsByTagName("t:Name")[0]?.textContent || 'Unknown';
                    }

                    messages.push({
                        sender, subject, body: cleanHtmlToText(body), received
                    });
                }

                messages.sort((a, b) => new Date(a.received) - new Date(b.received));
                resolve(messages);
            });
        });
    });
}

// ----------------------
// INIT OUTLOOK DATA
// ----------------------
function initOutlookData() {
    const item = Office.context.mailbox.item;
    if (!item) return;

    emailContext.meta = {
        senderName: item.sender?.displayName || 'Άγνωστος',
        senderEmail: item.sender?.emailAddress || '',
        subject: item.subject || '(Χωρίς θέμα)',
        receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('el-GR') : ''
    };

    startLoadingAnim(["Διαβάζω thread...", "Αναλύω συνομιλία...", "Φορτώνω ιστορικό..."]);

    // Προσπαθούμε με EWS πρώτα (Plan A)
    getFullConversationViaEWS()
        .then(messages => {
            emailContext.fullConversation = messages;
            const structured = messages.slice(-15).map(m => `FROM: ${m.sender}\nDATE: ${new Date(m.received).toLocaleString('el-GR')}\nSUBJECT: ${m.subject}\n\n${m.body}`).join("\n\n------------------------\n\n");
            
            if (structured.trim() !== "") {
                emailContext.text = structured;
                finishLoading();
            } else {
                fallbackCurrentMail(); // Αν το EWS γυρίσει κενό
            }
        })
        .catch(err => {
            console.warn("EWS Failed, switching to fallback:", err);
            fallbackCurrentMail(); // Plan B
        });
}

// ----------------------
// FALLBACK (Plan B - To Σίγουρο Text)
// ----------------------
function fallbackCurrentMail() {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) {
        stopLoadingAnim();
        return;
    }

    // ΑΛΛΑΓΗ ΕΔΩ: Γυρίσαμε στο CoercionType.Text που φέρνει όλο το σεντόνι 100%
    item.body.getAsync(Office.CoercionType.Text, (result) => {
        stopLoadingAnim();
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailContext.text = result.value;
            finishLoading();
        } else {
            console.error("Σφάλμα στο fallback");
        }
    });
}

// ----------------------
// LOADING & UI
// ----------------------
let loadingInterval;

function startLoadingAnim(messages) {
    const textEl = document.getElementById('loadingText');
    document.getElementById('loadingOverlay').style.display = 'flex';
    let i = 0;
    textEl.innerText = messages[0];
    loadingInterval = setInterval(() => {
        i = (i + 1) % messages.length;
        textEl.innerText = messages[i];
    }, 1400);
}

function stopLoadingAnim() {
    clearInterval(loadingInterval);
    document.getElementById('loadingOverlay').style.display = 'none';
}

function finishLoading() {
    stopLoadingAnim();
    if (config.autoSum && config.apiKey && emailContext.text && emailContext.text.length > 10) {
        generateSummary();
    }
}

// ----------------------
// SUMMARY
// ----------------------
async function generateSummary() {
    const prompt = `Κάνε μια Executive σύντομη σύνοψη (max 3-4 γραμμές) για το παρακάτω email thread. 
Περιέγραψε: Ποιος έστειλε το τελευταίο μήνυμα, τι ζητάει και αν υπάρχουν εκκρεμότητες.
THREAD:\n${emailContext.text}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
        });

        const data = await res.json();
        if (data.error) throw new Error(data.error.message);

        const summary = data.candidates[0].content.parts[0].text;
        document.getElementById('summaryText').innerText = summary;
    } catch (e) {
        console.error(e);
        document.getElementById('summaryText').innerText = 'Σφάλμα σύνοψης';
    }
}

// ----------------------
// QUICK ACTIONS
// ----------------------
function handleQuickAction(actionType) {
    if (!config.apiKey) { alert("Βάλε API Key"); return; }
    generateDraft(actionType, null);
}

// ----------------------
// TEXT SEND
// ----------------------
document.getElementById('sendTextBtn').onclick = () => {
    const txt = document.getElementById('textPrompt').value.trim();
    if (!txt) return;
    document.getElementById('textPrompt').value = '';
    generateDraft(txt, null);
};

// ----------------------
// TWEAK
// ----------------------
document.getElementById('tweakBtn').onclick = () => {
    const tweak = document.getElementById('tweakPrompt').value.trim();
    if (!tweak) return;
    const current = document.getElementById('draftTextarea').value;
    document.getElementById('tweakPrompt').value = '';
    
    // Force intent to draft for tweaks
    generateDraft(`Τροποποίησε το email έτσι: "${tweak}".\nΑΚΡΙΒΕΣ ΠΑΛΙΟ EMAIL:\n"${current}"`, null);
};

// ----------------------
// VOICE
// ----------------------
const voiceBtn = document.getElementById('voiceBtn');
const voiceStatus = document.getElementById('voiceStatus');

voiceBtn.onclick = () => {
    if (!config.apiKey) { alert("Βάλε API Key"); return; }
    if (isRecording) { stopRecording(); return; }

    if (Office.context.mailbox && Office.devicePermission) {
        Office.devicePermission.requestPermissionsAsync([Office.DevicePermissionType.microphone], (res) => {
            if (res.status === Office.AsyncResultStatus.Failed) {
                alert("Το μικρόφωνο απορρίφθηκε"); return;
            }
            if (res.value) window.location.reload(); 
            else startRecording();
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
            voiceStatus.innerText = 'Επεξεργασία ήχου...';
            const blob = new Blob(audioChunks, { type: mediaRecorder.mimeType });
            const reader = new FileReader();
            reader.readAsDataURL(blob);
            reader.onloadend = () => {
                const base64 = reader.result.split(',')[1];
                generateDraft('Αυτό είναι ηχητικό μήνυμα/εντολή.', { data: base64, mimeType: mediaRecorder.mimeType || 'audio/webm' });
            };
        };

        mediaRecorder.start();
        isRecording = true;
        voiceBtn.classList.remove('siri-idle');
        voiceBtn.classList.add('siri-listening');
        voiceBtn.innerHTML = `<i data-lucide="square" class="w-8 h-8 text-white"></i>`;
        lucide.createIcons();
        voiceStatus.innerText = 'Ακούω...';
    }).catch(err => {
        console.error(err);
        voiceStatus.innerText = 'Σφάλμα μικροφώνου';
    });
}

function stopRecording() {
    if (mediaRecorder && mediaRecorder.state !== 'inactive') mediaRecorder.stop();
    isRecording = false;
    voiceBtn.classList.remove('siri-listening');
    voiceBtn.classList.add('siri-idle');
    voiceBtn.innerHTML = `<i data-lucide="mic" class="w-8 h-8 opacity-70"></i>`;
    lucide.createIcons();
    voiceStatus.innerText = 'Κάντε κλικ για ομιλία';
}

// ----------------------
// GENERATE DRAFT (THE ENGINE)
// ----------------------
async function generateDraft(instruction, audioObj) {
    voiceStatus.innerText = 'Σκέφτομαι...';

    const systemPrompt = `Είσαι Executive AI Assistant. Διάβασε το παρακάτω ιστορικό email για να έχεις πλήρες context.

ΣΤΟΙΧΕΙΑ: 
Από: ${emailContext.meta.senderName} (${emailContext.meta.senderEmail})
Θέμα: ${emailContext.meta.subject}

EMAIL THREAD:
"""
${emailContext.text}
"""

USER REQUEST (Εντολή Αφεντικού):
${instruction}

ΚΑΘΗΚΟΝ:
Κατάλαβε αν ο χρήστης ζητάει να ΓΡΑΨΕΙΣ ΕΝΑ EMAIL ("απάντα", "στείλε", "ευχαρίστησε") ή αν σου κάνει μια ΕΡΩΤΗΣΗ ("ποιος είναι;", "τι έγινε;", "πότε λήγει;").

Απάντησε ΑΥΣΤΗΡΑ ΚΑΙ ΜΟΝΟ με ένα JSON αντικείμενο, χωρίς markdown:
{
 "intent": "draft" ή "question",
 "content": "Το παραγόμενο κείμενο εδώ"
}

ΚΑΝΟΝΕΣ ΑΝ intent == "draft":
- Γράψε ΑΠΕΥΘΕΙΑΣ το κείμενο του email απαντώντας στο ΤΕΛΕΥΤΑΙΟ μήνυμα του πελάτη.
- ΤΟΝΟΣ: ${config.tone}. EXTRA ΚΑΝΟΝΕΣ: ${config.customPrompt}.
- Απαγορεύεται να κάνεις υποθέσεις, να επινοείς τιμές ή ημερομηνίες που δεν ειπώθηκαν.

ΚΑΝΟΝΕΣ ΑΝ intent == "question":
- Δώσε μια απλή, άμεση και κατανοητή απάντηση στον χρήστη. ΜΗΝ τη γράφεις σαν email.`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const parts = [{ text: systemPrompt }];

        if (audioObj) {
            parts.push({ inlineData: { mimeType: audioObj.mimeType, data: audioObj.data } });
        }

        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts }],
                generationConfig: { responseMimeType: "application/json" }
            })
        });

        const data = await res.json();
        if (data.error) throw new Error(data.error.message);

        const raw = data.candidates[0].content.parts[0].text.trim();
        const parsed = JSON.parse(raw.replace(/```json/gi, '').replace(/```/g, '').trim());

        if (parsed.intent === 'question') {
            document.getElementById('answerText').innerText = parsed.content;
            navigate('view-answer');
        } else {
            document.getElementById('draftTextarea').value = parsed.content;
            navigate('view-draft');
        }

        voiceStatus.innerText = 'Κάντε κλικ για ομιλία';

    } catch (e) {
        console.error(e);
        alert("Σφάλμα AI. Η απάντηση ίσως δεν ήταν έγκυρο JSON.");
        voiceStatus.innerText = 'Σφάλμα';
    }
}

// ----------------------
// INSERT TO OUTLOOK
// ----------------------
document.getElementById('insertOutlookBtn').onclick = () => {
    const finalTxt = document.getElementById('draftTextarea').value;
    Office.context.mailbox.item.displayReplyForm(finalTxt);
    document.getElementById('draftTextarea').value = '';
    navigate('view-main');
};

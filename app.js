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

function navigate(viewId) {
    document.querySelectorAll('.view-section').forEach(el => {
        el.classList.remove('view-active');
    });
    document.getElementById(viewId).classList.add('view-active');
}

// ----------------------
// CLEAN HTML
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
        .replace(/&nbsp;/g, ' ')
        .trim();
}

// ----------------------
// GET FULL THREAD VIA REST API (ΤΟ ΑΠΟΛΥΤΟ ΚΟΛΠΟ)
// ----------------------
async function getFullConversationViaREST() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        const convId = item.conversationId;

        if (!convId) {
            reject(new Error("Δεν βρέθηκε Conversation ID"));
            return;
        }

        // Ζητάμε Token Πρόσβασης για το REST API
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("Το Token απέτυχε. Μήπως λείπει το ReadWriteMailbox στο Manifest;"));
                return;
            }

            const token = result.value;
            const restUrl = Office.context.mailbox.restUrl;
            
            if (!restUrl) {
                reject(new Error("Δεν υπάρχει REST URL"));
                return;
            }

            // Τραβάμε ΟΛΑ τα μηνύματα της συνομιλίας χρονολογικά
            const url = `${restUrl}/v2.0/me/messages?$filter=ConversationId eq '${convId}'&$select=Sender,Subject,Body,DateTimeReceived&$orderby=DateTimeReceived asc&$top=20`;

            fetch(url, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json'
                }
            })
            .then(res => {
                if (!res.ok) throw new Error(`REST Error: ${res.status}`);
                return res.json();
            })
            .then(data => {
                if (data && data.value && data.value.length > 0) {
                    const messages = data.value.map(m => {
                        let senderName = "Άγνωστος";
                        if (m.Sender && m.Sender.EmailAddress) {
                            senderName = m.Sender.EmailAddress.Name || m.Sender.EmailAddress.Address;
                        }
                        return {
                            sender: senderName,
                            subject: m.Subject || "(Χωρίς Θέμα)",
                            body: cleanHtmlToText(m.Body.Content),
                            received: m.DateTimeReceived
                        };
                    });
                    resolve(messages);
                } else {
                    reject(new Error("Η συνομιλία βρέθηκε κενή στο REST"));
                }
            })
            .catch(err => reject(err));
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

    startLoadingAnim([
        "Διαβάζω thread...",
        "Αναλύω συνομιλία...",
        "Φορτώνω ιστορικό..."
    ]);

    // Προσπάθεια να φέρει όλο το Thread μέσω REST API
    getFullConversationViaREST()
        .then(messages => {
            emailContext.fullConversation = messages;
            
            // Ενώνουμε τα μηνύματα με σαφή διαχωριστικά για το AI
            const structured = messages
                .map(m => `ΑΠΟ: ${m.sender}\nΗΜΕΡΟΜΗΝΙΑ: ${new Date(m.received).toLocaleString('el-GR')}\nΜΗΝΥΜΑ:\n${m.body}`)
                .join("\n\n=== ΤΕΛΟΣ ΜΗΝΥΜΑΤΟΣ ===\n\n");

            emailContext.text = structured;
            console.log("Επιτυχία REST API. Μήκος:", emailContext.text.length);
            finishLoading();
        })
        .catch(err => {
            console.warn("REST API Failed (Μήπως δεν έβαλες ReadWriteMailbox;):", err);
            fallbackCurrentMail();
        });
}

// ----------------------
// FALLBACK (Plan B)
// ----------------------
function fallbackCurrentMail() {
    const item = Office.context.mailbox.item;
    if (!item.body) {
        stopLoadingAnim();
        return;
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
        stopLoadingAnim();
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailContext.text = result.value;
            finishLoading();
        }
    });
}

// ----------------------
// LOADING ANIMATION
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
    if (config.autoSum && config.apiKey && emailContext.text) {
        generateSummary();
    }
}

// ----------------------
// SUMMARY
// ----------------------
async function generateSummary() {
    const prompt = `Κάνε σύντομη executive σύνοψη του παρακάτω email thread (max 4 γραμμές).
Περιέγραψε: Ποιος έστειλε το τελευταίο μήνυμα, τι ζητάει, και αν υπάρχουν εκκρεμότητες.

THREAD:
${emailContext.text}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const res = await fetch(url, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
        });

        const data = await res.json();
        if (data.error) throw new Error(data.error.message);

        document.getElementById('summaryText').innerText = data.candidates[0].content.parts[0].text;
    } catch (e) {
        console.error(e);
        document.getElementById('summaryText').innerText = 'Σφάλμα σύνοψης';
    }
}

// ----------------------
// QUICK ACTIONS & TWEAKS
// ----------------------
function handleQuickAction(actionType) {
    if (!config.apiKey) { alert("Βάλε API Key"); return; }
    generateDraft(actionType, null);
}

document.getElementById('sendTextBtn').onclick = () => {
    const txt = document.getElementById('textPrompt').value.trim();
    if (!txt) return;
    document.getElementById('textPrompt').value = '';
    generateDraft(txt, null);
};

document.getElementById('tweakBtn').onclick = () => {
    const tweak = document.getElementById('tweakPrompt').value.trim();
    if (!tweak) return;
    const current = document.getElementById('draftTextarea').value;
    document.getElementById('tweakPrompt').value = '';
    generateDraft(`Τροποποίησε το προηγούμενο draft email σου σύμφωνα με αυτή την οδηγία: ${tweak}\n\nΠΑΛΙΟ EMAIL:\n${current}`, null);
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
                generateDraft('Αυτό είναι το ηχητικό μήνυμα/εντολή του χρήστη.', { data: base64, mimeType: mediaRecorder.mimeType || 'audio/webm' });
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
// GENERATE DRAFT (AI ENGINE)
// ----------------------
async function generateDraft(instruction, audioObj) {
    voiceStatus.innerText = 'Σκέφτομαι...';

    const systemPrompt = `Είσαι Executive AI Assistant. Έχεις το πλήρες ιστορικό μιας συνομιλίας (Email Thread).
Τα μηνύματα είναι χωρισμένα με "=== ΤΕΛΟΣ ΜΗΝΥΜΑΤΟΣ ===". Το τελευταίο χρονολογικά μήνυμα βρίσκεται στο κάτω μέρος της λίστας.

EMAIL THREAD:
${emailContext.text}

USER REQUEST:
${instruction}

ΚΑΘΗΚΟΝ:
Κατάλαβε αν ο χρήστης σου ζητάει να ΓΡΑΨΕΙΣ ΕΝΑ EMAIL (π.χ. "απάντα", "ευχαρίστησέ τον") ή αν σου κάνει μια ΕΡΩΤΗΣΗ/ΣΥΖΗΤΗΣΗ (π.χ. "τι λέει εδώ;", "ποιος είναι;").
Απάντησε ΑΥΣΤΗΡΑ με JSON:
{
 "intent": "draft" ή "question",
 "content": "Το κείμενό σου"
}

ΚΑΝΟΝΕΣ:
Αν intent="draft": Γράψε έτοιμο email απαντώντας στο ΤΕΛΕΥΤΑΙΟ μήνυμα του thread. Τόνος: ${config.tone}.
Αν intent="question": Απάντα ξεκάθαρα στην ερώτηση, μην το γράψεις σαν email.`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const parts = [{ text: systemPrompt }];
        if (audioObj) parts.push({ inlineData: { mimeType: audioObj.mimeType, data: audioObj.data } });

        const res = await fetch(url, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts }],
                generationConfig: { responseMimeType: "application/json" }
            })
        });

        const data = await res.json();
        if (data.error) throw new Error(data.error.message);

        const raw = data.candidates[0].content.parts[0].text.trim();
        const parsed = JSON.parse(raw.replace(/```json/g, '').replace(/```/g, '').trim());

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
        alert("AI Error. Βεβαιώσου ότι ο ήχος ήταν καθαρός.");
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

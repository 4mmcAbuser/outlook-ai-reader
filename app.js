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
let emailContext = { text: '', meta: {} };
let isRecording = false;
let mediaRecorder;
let audioChunks = [];

// --- INITIALIZATION ---
Office.onReady((info) => {
    loadSettings();
    if (info.host === Office.HostType.Outlook) {
        initOutlookData();
    }
});

function loadSettings() {
    const saved = localStorage.getItem('aiAssistConfig');
    if (saved) config = { ...config, ...JSON.parse(saved) };
    
    // Populate settings view
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

// --- OUTLOOK DATA FETCH ---
function initOutlookData() {
    const item = Office.context.mailbox.item;
    emailContext.meta = {
        senderName: item.sender ? item.sender.displayName : "Άγνωστος",
        senderEmail: item.sender ? item.sender.emailAddress : "unknown@mail.com",
        subject: item.subject,
        toRecipients: item.to ? item.to.map(r => r.displayName).join(", ") : "Μόνο εγώ"
    };

    item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailContext.text = result.value;
            if (config.autoSum && config.apiKey) {
                generateSummary();
            } else {
                showManualSummaryBtn();
            }
        }
    });
}

// --- FAKE LOADING ANIMATION ---
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
    document.getElementById('summaryText').innerText = "Η αυτόματη σύνοψη είναι ανενεργή.";
    const btn = document.getElementById('manualSummaryBtn');
    btn.style.display = 'block';
    btn.onclick = () => { btn.style.display = 'none'; generateSummary(); };
}

// --- SUMMARY GENERATION ---
async function generateSummary() {
    if (!config.apiKey) return;
    startLoadingAnim(["Διαβάζω το email...", "Αναλύω τα δεδομένα...", "Ετοιμάζω σύνοψη..."]);
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${config.apiKey}`;
    const prompt = `Κάνε μια πολύ σύντομη, περιεκτική σύνοψη (max 3-4 γραμμές) στα Ελληνικά για το παρακάτω email. Ποιος στέλνει και τι ζητάει.\nEmail: ${emailContext.text}`;
    
    try {
        const res = await fetch(url, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
        });
        const data = await res.json();
        stopLoadingAnim();
        
        const summary = data.candidates[0].content.parts[0].text;
        const sumEl = document.getElementById('summaryText');
        sumEl.innerText = summary;
        
        // Check if expand button is needed
        if (sumEl.scrollHeight > 96) { // 96px is max-h-24
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
        document.getElementById('summaryText').innerText = "Σφάλμα κατά τη σύνοψη.";
    }
}

// --- ACTION LOGIC (Text, Quick Actions, Voice) ---

// 1. Quick Actions
function handleQuickAction(actionType) {
    if(!config.apiKey) return alert("Βάλε API Key στις ρυθμίσεις");
    generateDraft(actionType, null);
}

// 2. Text Input
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
        generateDraft(`Τροποποίησε το κείμενο: "${currentDraft}". ΟΔΗΓΙΑ ΤΡΟΠΟΠΟΙΗΣΗΣ: ${tweak}`, null);
    }
}

// 3. Voice Input
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
                generateDraft("Απάντα με βάση την ηχητική εντολή.", { data: b64, mimeType: mime });
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

// --- CORE GENERATOR ENGINE ---
async function generateDraft(instruction, audioObj) {
    voiceStatus.innerText = "Δημιουργία απάντησης...";
    
    // Construct the Anti-Hallucination & Tone Prompt
    const systemPrompt = `Είσαι Executive Assistant. 
Ανάλυσε το email και γράψε ΑΠΕΥΘΕΙΑΣ το κείμενο της απάντησης (χωρίς JSON, χωρίς χαιρετισμούς δικούς σου προς εμένα).
ΤΟΝΟΣ ΑΠΑΝΤΗΣΗΣ: ${config.tone}
ΕΞΤΡΑ ΚΑΝΟΝΕΣ: ${config.customPrompt}
ANTI-HALLUCINATION ΚΑΝΟΝΕΣ: Μην κάνεις υποθέσεις. Μην προσθέτεις ονόματα, ημερομηνίες ή ποσά που δεν αναφέρθηκαν στο email ή στην εντολή. Βασίσου αυστηρά στα δεδομένα.

ΣΤΟΙΧΕΙΑ ΠΕΛΑΤΗ: Αποστολέας: ${emailContext.meta.senderName}, Θέμα: ${emailContext.meta.subject}
ΙΣΤΟΡΙΚΟ EMAIL: "${emailContext.text}"

ΕΝΤΟΛΗ ΧΡΗΣΤΗ: ${instruction}`;

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
    
    let parts = [{ text: systemPrompt }];
    if (audioObj) {
        parts.push({ inlineData: { mimeType: audioObj.mimeType, data: audioObj.data } });
    }

    try {
        const res = await fetch(url, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ contents: [{ parts: parts }] })
        });
        
        const data = await res.json();
        if(data.error) throw new Error(data.error.message);
        
        const draft = data.candidates[0].content.parts[0].text.trim();
        
        // Show Draft Page
        document.getElementById('draftTextarea').value = draft;
        navigate('view-draft');
        voiceStatus.innerText = "Κάντε κλικ για ομιλία";
        
    } catch (e) {
        alert("Σφάλμα AI: " + e.message);
        voiceStatus.innerText = "Σφάλμα. Προσπαθήστε ξανά.";
    }
}

// --- INSERT TO OUTLOOK ---
document.getElementById('insertOutlookBtn').onclick = () => {
    const finalTxt = document.getElementById('draftTextarea').value;
    Office.context.mailbox.item.displayReplyForm(finalTxt);
    // Return to main and clear
    document.getElementById('draftTextarea').value = '';
    navigate('view-main');
};

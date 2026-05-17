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

    // ΑΛΛΑΓΗ 1: Χρησιμοποιούμε Html αντί για Text για να διαβάζει όλο το ιστορικό της συνομιλίας
    item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            
            // ΑΛΛΑΓΗ 2: Καθαρίζουμε το HTML για να το κάνουμε απλό κείμενο, διατηρώντας τα παλιά μηνύματα
            const parser = new DOMParser();
            const doc = parser.parseFromString(result.value, 'text/html');
            emailContext.text = doc.body.textContent || doc.body.innerText || "";
            
            if (config.autoSum && config.apiKey) {
                generateSummary();
            } else {
                showManualSummaryBtn();
            }
        } else {
             console.error("Σφάλμα ανάγνωσης email από το Outlook.");
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
    
    // Χρησιμοποιούμε το επιλεγμένο μοντέλο και για τη σύνοψη (ώστε αν το 1.5 flash-latest έχει θέμα, να παίρνει το lite)
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
    const prompt = `Κάνε μια πολύ σύντομη, περιεκτική σύνοψη (max 3-4 γραμμές) στα Ελληνικά για την παρακάτω συνομιλία email. Ποιος στέλνει το τελευταίο μήνυμα και τι ζητάει.\nΣυνομιλία:\n${emailContext.text}`;
    
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
        document.getElementById('summaryText').innerText = "Σφάλμα κατά τη σύνοψη: " + e.message;
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
    
    // ΑΛΛΑΓΗ 3: Το νέο System Prompt που εκπαιδεύει το AI να διαβάζει τη συνομιλία χρονικά (από κάτω προς τα πάνω)
    const systemPrompt = `Είσαι ένας έμπειρος Executive Assistant. Σου δίνεται ένα ολόκληρο ιστορικό συνομιλίας (Email Thread) το οποίο περιέχει πολλαπλά μηνύματα.

ΚΑΝΟΝΕΣ ΑΝΑΛΥΣΗΣ ΙΣΤΟΡΙΚΟΥ:
1. Το κείμενο περιέχει όλη την κουβέντα. Τα πιο παλιά μηνύματα είναι συνήθως στο κάτω μέρος (συχνά ξεκινούν με γραμμές όπως "Στις Σάβ 16 Μαΐ... έγραψε:") και τα πιο πρόσφατα είναι στην κορυφή.
2. Διάβασε προσεκτικά ΟΛΑ τα μηνύματα για να καταλάβεις τη ροή της κουβέντας, τι έχει συμφωνηθεί, τι ρώτησε ο πελάτης πριν και τι του απαντήσαμε.
3. Εντόπισε το ΤΕΛΕΥΤΑΙΟ ΚΑΙ ΠΙΟ ΠΡΟΣΦΑΤΟ μήνυμα που στάλθηκε από τον πελάτη (${emailContext.meta.senderName}). 
4. Γράψε ΑΠΕΥΘΕΙΑΣ το κείμενο της απάντησης (χωρίς JSON) απαντώντας ΑΠΟΚΛΕΙΣΤΙΚΑ σε αυτό το τελευταίο μήνυμα, αλλά έχοντας ως γνώμονα όσα ειπώθηκαν παραπάνω.

ANTI-HALLUCINATION & ΤΟΝΟΣ:
- ΤΟΝΟΣ ΑΠΑΝΤΗΣΗΣ: ${config.tone}
- ΕΞΤΡΑ ΚΑΝΟΝΕΣ ΧΡΗΣΤΗ: ${config.customPrompt}
- Μην κάνεις υποθέσεις. Μην επινοείς στοιχεία. Αν ο πελάτης ρωτάει "πόσο θα κοστίσει" ή "πόσες μέρες θα χρειαστούν", άκουσε την εντολή του αφεντικού στον ήχο/κείμενο. Αν δεν αναφέρεται συγκεκριμένη πληροφορία, γράψε στην απάντηση ότι "θα το εξετάσουμε και θα σας ενημερώσουμε άμεσα". ΜΗΝ βγάλεις στην τύχη σου νούμερα.

ΣΤΟΙΧΕΙΑ ΜΗΝΥΜΑΤΟΣ: 
Αποστολέας: ${emailContext.meta.senderName} (${emailContext.meta.senderEmail})
Θέμα: ${emailContext.meta.subject}

ΟΛΟΚΛΗΡΟ ΤΟ ΙΣΤΟΡΙΚΟ ΣΥΝΟΜΙΛΙΑΣ (THREAD):
"""
${emailContext.text}
"""

ΟΔΗΓΙΑ ΑΦΕΝΤΙΚΟΥ (Από ήχο ή κείμενο): ${instruction}`;

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

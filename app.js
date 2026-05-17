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

// --- OUTLOOK DATA FETCH (THE HTML FIX) ---
function initOutlookData() {
    const item = Office.context.mailbox.item;
    emailContext.meta = {
        senderName: item.sender ? item.sender.displayName : "Άγνωστος",
        senderEmail: item.sender ? item.sender.emailAddress : "unknown@mail.com",
        subject: item.subject,
        toRecipients: item.to ? item.to.map(r => r.displayName).join(", ") : "Μόνο εγώ"
    };

    // Χρησιμοποιούμε εναλλακτικά το CoercionType.Html αλλά με πλήρη υποστήριξη thread
    if (item.body) {
        item.body.getAsync(Office.CoercionType.Html, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                let rawHtml = result.value;
                
                // Αν το HTML περιέχει tags τύπου 'divRplyFwdMsg' (κλασικό Outlook thread marker)
                // σημαίνει ότι έχει έρθει όλο το ιστορικό. Αν όχι, το Outlook μας περιόρισε.
                let cleanText = rawHtml
                    .replace(/<style[^>]*>.*?<\/style>/gi, '')
                    .replace(/<script[^>]*>.*?<\/script>/gi, '')
                    .replace(/<br\s*[\/]?>/gi, '\n')
                    .replace(/<\/p>/gi, '\n\n')
                    .replace(/<\/div>/gi, '\n')
                    .replace(/<[^>]+>/g, '');
                
                const txt = document.createElement('textarea');
                txt.innerHTML = cleanText;
                emailContext.text = txt.value.trim();
                
                if (config.autoSum && config.apiKey) {
                    generateSummary();
                } else {
                    showManualSummaryBtn();
                }
            } else {
                showManualSummaryBtn();
            }
        });
    }
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
    startLoadingAnim(["Διαβάζω το Thread...", "Αναλύω τη συνομιλία...", "Ετοιμάζω σύνοψη..."]);
    
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
        // Force the intent to be a draft when tweaking
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
    
    // ΝΕΟ PROMPT: Του λέμε να ξεχωρίζει τι του ζητάς και να το επιστρέφει σε JSON
    const systemPrompt = `Είσαι ένας έμπειρος Executive Assistant. Έχεις μπροστά σου το ιστορικό μιας συνομιλίας (Email Thread). 
Τα πιο πρόσφατα μηνύματα είναι στην κορυφή, τα παλαιότερα στο κάτω μέρος.

ΣΤΟΙΧΕΙΑ ΜΗΝΥΜΑΤΟΣ: 
Αποστολέας: ${emailContext.meta.senderName} (${emailContext.meta.senderEmail})
Θέμα: ${emailContext.meta.subject}

ΙΣΤΟΡΙΚΟ ΣΥΝΟΜΙΛΙΑΣ:
"""
${emailContext.text}
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
        
        // Clean up markdown just in case
        const cleanJsonString = rawResponse.replace(/```json/gi, '').replace(/```/g, '').trim();
        const parsedResponse = JSON.parse(cleanJsonString);
        
        // --- SMART ROUTING BASED ON INTENT ---
        if (parsedResponse.intent === "question") {
            // Ο χρήστης έκανε ερώτηση -> Πάμε στην οθόνη ερωτήσεων
            document.getElementById('answerText').innerText = parsedResponse.content;
            navigate('view-answer');
        } else {
            // Ο χρήστης ζήτησε email -> Πάμε στο Draft
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

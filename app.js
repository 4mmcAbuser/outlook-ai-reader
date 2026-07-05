// Initialize Lucide Icons
lucide.createIcons();

// ----------------------
// STATE MANAGEMENT
// ----------------------
let config = {
    apiKey: '',
    model: 'gemma-4-31b-it',
    tone: 'Επαγγελματικός, ευγενικός και σοβαρός.',
    autoSum: true,
    customPrompt: '',
    agentMemory: '' // Agent Memory state added
};

let emailContext = {
    text: '',
    meta: {},
    fullConversation: []
};

let isRecording = false;
let mediaRecorder;
let audioChunks = [];
let currentSpeechUtterance = null; // TTS tracker

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
                console.log("🔄 Item changed - reloading context");
                if (window.speechSynthesis) window.speechSynthesis.cancel(); // Stop reading if mail changes
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

    document.getElementById('setApiKey').value = config.apiKey || '';
    document.getElementById('setModel').value = config.model;
    document.getElementById('setTone').value = config.tone;
    document.getElementById('setAutoSum').checked = config.autoSum;
    document.getElementById('setCustomPrompt').value = config.customPrompt || '';
    document.getElementById('setAgentMemory').value = config.agentMemory || '';

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
    config.agentMemory = document.getElementById('setAgentMemory').value.trim();

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
// UTILITIES
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
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .trim();
}

// ----------------------
// TEXT TO SPEECH (FOR CLIENT'S HEADPHONES)
// ----------------------
function speakSummary() {
    if (!window.speechSynthesis) {
        alert("Το σύστημά σας δεν υποστηρίζει Text-to-Speech.");
        return;
    }

    // Toggle speech if already talking
    if (window.speechSynthesis.speaking) {
        window.speechSynthesis.cancel();
        document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Ακρόαση`;
        lucide.createIcons();
        return;
    }

    const textToRead = document.getElementById('summaryText').innerText;
    if (textToRead === 'Περιμένω ανάλυση...' || textToRead.startsWith('⚠️')) return;

    currentSpeechUtterance = new SpeechSynthesisUtterance(textToRead);
    currentSpeechUtterance.lang = 'el-GR'; // Force Greek formatting
    currentSpeechUtterance.rate = 1.0;

    currentSpeechUtterance.onend = () => {
        document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Ακρόαση`;
        lucide.createIcons();
    };

    document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-x" class="w-3.5 h-3.5 text-red-400"></i> Διακοπή`;
    lucide.createIcons();
    
    window.speechSynthesis.speak(currentSpeechUtterance);
}

// ----------------------
// FETCHING METHOD VIA OUTLOOK REST API
// ----------------------
async function getFullConversationViaREST() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        const convId = item.conversationId;

        if (!convId) {
            reject(new Error("Δεν βρέθηκε Conversation ID"));
            return;
        }

        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (tokenResult) => {
            if (tokenResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("Αποτυχία λήψης token."));
                return;
            }

            const token = tokenResult.value;
            const restUrl = Office.context.mailbox.restUrl;
            
            if (!restUrl) {
                reject(new Error("Δεν βρέθηκε REST URL"));
                return;
            }

            const url = `${restUrl}/v2.0/me/messages?$filter=ConversationId eq '${convId}'&$select=Sender,Subject,Body,DateTimeReceived,ConversationIndex&$orderby=DateTimeReceived asc&$top=50`;

            fetch(url, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json',
                    'prefer': 'outlook.body-content-type="text"'
                }
            })
            .then(res => res.ok ? res.json() : reject(new Error("REST fetch failed")))
            .then(data => {
                if (data && data.value && data.value.length > 0) {
                    const messages = data.value
                        .filter(m => m.Body && m.Body.Content)
                        .map(m => {
                            let senderName = m.Sender?.EmailAddress?.Name || "Άγνωστος";
                            let senderEmail = m.Sender?.EmailAddress?.Address || "";
                            const bodyContent = m.Body.ContentType === 'HTML' ? cleanHtmlToText(m.Body.Content) : m.Body.Content;
                            
                            return {
                                sender: `${senderName} <${senderEmail}>`,
                                subject: m.Subject || "(Χωρίς Θέμα)",
                                body: bodyContent,
                                received: m.DateTimeReceived
                            };
                        });
                    resolve(messages);
                } else {
                    reject(new Error("Empty thread"));
                }
            })
            .catch(err => reject(err));
        });
    });
}

function initOutlookData() {
    const item = Office.context.mailbox.item;
    if (!item) return;

    emailContext.meta = {
        senderName: item.sender?.displayName || 'Άγνωστος',
        senderEmail: item.sender?.emailAddress || '',
        subject: item.subject || '(Χωρίς θέμα)',
        receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('el-GR') : ''
    };

    startLoadingAnim(["📡 Σύνδεση...", "🔍 Ανάλυση Thread...", "🤖 Κατηγοριοποίηση..."]);

    getFullConversationViaREST()
        .then(messages => {
            emailContext.fullConversation = messages;
            const structured = messages.map((m, idx) => `--- ΜΗΝΥΜΑ #${idx + 1} ---\nΑΠΟΣΤΟΛΕΑΣ: ${m.sender}\nΘΕΜΑ: ${m.subject}\nΠΕΡΙΕΧΟΜΕΝΟ:\n${m.body}\n`).join("\n");
            emailContext.text = structured;
            finishLoading();
        })
        .catch(() => {
            fallbackCurrentMail();
        });
}

function fallbackCurrentMail() {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) {
        stopLoadingAnim();
        emailContext.text = "Σφάλμα ανάγνωσης.";
        finishLoading();
        return;
    }
    item.body.getAsync(Office.CoercionType.Text, (result) => {
        stopLoadingAnim();
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailContext.text = `--- ΤΡΕΧΟΝ ΜΗΝΥΜΑ ---\nΑΠΟΣΤΟΛΕΑΣ: ${emailContext.meta.senderName}\nΘΕΜΑ: ${emailContext.meta.subject}\n\n${result.value}`;
            finishLoading();
        }
    });
}

// ----------------------
// LOADING MANAGEMENT
// ----------------------
let loadingInterval;
function startLoadingAnim(messages) {
    const textEl = document.getElementById('loadingText');
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = 'flex';
        let i = 0; textEl.innerText = messages[0];
        loadingInterval = setInterval(() => {
            i = (i + 1) % messages.length;
            textEl.innerText = messages[i];
        }, 1200);
    }
}
function stopLoadingAnim() {
    if (loadingInterval) clearInterval(loadingInterval);
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.style.display = 'none';
}
function finishLoading() {
    stopLoadingAnim();
    if (config.autoSum && config.apiKey && emailContext.text) {
        generateSummaryAndAudit();
    }
}

// ----------------------
// EMAIL AUDIT & SUMMARY ENGINE (DUAL INTERFACE)
// ----------------------
async function generateSummaryAndAudit() {
    if (!config.apiKey) {
        document.getElementById('summaryText').innerText = '⚠️ Παρακαλώ προσθέστε API Key στις ρυθμίσεις.';
        return;
    }

    const prompt = `Ανάλυσε το παρακάτω email thread.
Επίστρεψε ΑΥΣΤΗΡΑ ένα αντικείμενο JSON (χωρίς markdown κώδικα) με τα εξής πεδία:
{
  "summary": "Μια σύντομη executive σύνοψη (max 4 γραμμές) στα Ελληνικά.",
  "category": "High Priority" ή "Internal" ή "Newsletter" ή "Spam"
}

THREAD:
${emailContext.text.substring(0, 8000)}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { 
                    temperature: 0.2 
                }
            })
        });

        // 1. Έλεγχος αν η HTTP απάντηση είναι επιτυχής (π.χ. 200 OK)
        if (!res.ok) {
            const errBody = await res.json().catch(() => ({}));
            throw new Error(`HTTP ${res.status}: ${errBody.error?.message || res.statusText}`);
        }

        const data = await res.json();
        
        // 2. Έλεγχος αν το API επέστρεψε payload σφάλματος
        if (data.error) {
            throw new Error(`API Error: ${data.error.message}`);
        }

        let raw = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        if (!raw) {
            throw new Error("Το μοντέλο επέστρεψε κενή απάντηση ή μπλοκαρίστηκε από τα φίλτρα ασφαλείας.");
        }

        // Καθαρισμός markdown σε περίπτωση που το μοντέλο έβαλε ```json ... ```
        raw = raw.replace(/```json/gi, '').replace(/```/gi, '').trim();
        
        // 3. Προσπάθεια Parsing του JSON
        let resObj;
        try {
            resObj = JSON.parse(raw);
        } catch (jsonErr) {
            throw new Error("Το μοντέλο δεν επέστρεψε έγκυρη δομή JSON. Απάντηση: " + raw.substring(0, 100));
        }
        
        if (!resObj.summary) {
            throw new Error("Το JSON που παρήχθη δεν περιέχει το πεδίο 'summary'.");
        }
        
        // Ενημέρωση του UI με τη σύνοψη
        document.getElementById('summaryText').innerText = resObj.summary;
        
        // Ενημέρωση των Analytics & του Category Badge
        updateAuditMetrics(resObj.category || 'Spam');
        renderCategoryBadge(resObj.category || 'Spam');

        if (resObj.summary.length > 150) {
            document.getElementById('summaryContent').classList.add('max-h-24');
            document.getElementById('summaryFade')?.classList.remove('hidden');
            document.getElementById('expandSummaryBtn')?.classList.remove('hidden');
        }
    } catch (e) {
        // Καταγραφή στο Console του προγράμματος περιήγησης
        console.error("❌ Summary Generation Error:", e);
        // Εμφάνιση του πραγματικού σφάλματος στον χρήστη για εύκολο debugging
        document.getElementById('summaryText').innerText = `⚠️ Σφάλμα: ${e.message}`;
    }
}

// Render dynamic badges based on categories
function renderCategoryBadge(cat) {
    const badge = document.getElementById('emailCategoryBadge');
    badge.className = "text-[10px] font-bold px-2 py-0.5 rounded-full transition-all duration-300 ";
    
    switch(cat) {
        case 'High Priority':
            badge.innerText = '🔥 Υψηλή Προτεραιότητα';
            badge.classList.add('bg-red-500/20', 'text-red-400');
            break;
        case 'Internal':
            badge.innerText = '💼 Εσωτερικό / Εταιρικό';
            badge.classList.add('bg-blue-500/20', 'text-blue-400');
            break;
        case 'Newsletter':
            badge.innerText = '📢 Newsletter';
            badge.classList.add('bg-yellow-500/20', 'text-yellow-400');
            break;
        default:
            badge.innerText = '🗑️ Χαμηλή Σημασία';
            badge.classList.add('bg-zinc-800', 'text-zinc-400');
    }
}

// Log dynamic data updates locally inside storage
function updateAuditMetrics(cat) {
    let auditData = JSON.parse(localStorage.getItem('emailAuditStore')) || { total: 0, high: 0, internal: 0, news: 0, spam: 0 };
    
    auditData.total += 1;
    if (cat === 'High Priority') auditData.high += 1;
    else if (cat === 'Internal') auditData.internal += 1;
    else if (cat === 'Newsletter') auditData.news += 1;
    else auditData.spam += 1;

    localStorage.setItem('emailAuditStore', JSON.stringify(auditData));
}

// ----------------------
// AUDIT DASHBOARD RENDERING
// ----------------------
function openAuditDashboard() {
    navigate('view-audit');
    const data = JSON.parse(localStorage.getItem('emailAuditStore')) || { total: 0, high: 0, internal: 0, news: 0, spam: 0 };
    
    document.getElementById('auditTotal').innerText = data.total;
    // Simple math metric estimation: 3 mins saved per processed email setup
    document.getElementById('auditSavedTime').innerText = (data.total * 3) + 'λ';

    const calcPct = (val) => data.total > 0 ? Math.round((val / data.total) * 100) : 0;

    const pHigh = calcPct(data.high), pInternal = calcPct(data.internal), pNews = calcPct(data.news), pSpam = calcPct(data.spam);

    document.getElementById('pct-high').innerText = pHigh + '%';
    document.getElementById('bar-high').style.width = pHigh + '%';

    document.getElementById('pct-internal').innerText = pInternal + '%';
    document.getElementById('bar-internal').style.width = pInternal + '%';

    document.getElementById('pct-news').innerText = pNews + '%';
    document.getElementById('bar-news').style.width = pNews + '%';

    document.getElementById('pct-spam').innerText = pSpam + '%';
    document.getElementById('bar-spam').style.width = pSpam + '%';
}

function clearAuditData() {
    localStorage.removeItem('emailAuditStore');
    openAuditDashboard();
}

// Expand text action UI toggles
document.getElementById('expandSummaryBtn')?.addEventListener('click', function() {
    const summaryEl = document.getElementById('summaryContent');
    const fadeEl = document.getElementById('summaryFade');
    if (summaryEl.classList.contains('max-h-24')) {
        summaryEl.classList.remove('max-h-24');
        fadeEl?.classList.add('hidden');
        this.innerText = 'Λιγότερα...';
    } else {
        summaryEl.classList.add('max-h-24');
        fadeEl?.classList.remove('hidden');
        this.innerText = 'Περισσότερα...';
    }
});

// ----------------------
// GENERATE DRAFT WITH AGENTIC MEMORY
// ----------------------
async function generateDraft(instruction, audioObj) {
    if (!config.apiKey) { navigate('view-settings'); return; }
    voiceStatus.innerText = '🤖 Σκέφτομαι...';

    const systemPrompt = `Είσαι Executive AI Assistant για επαγγελματική αλληλογραφία.
Έχεις πρόσβαση στο ΠΛΗΡΕΣ ιστορικό μιας συνομιλίας (Email Thread).

📧 EMAIL THREAD:
${emailContext.text.substring(0, 20000)}

👤 USER INSTRUCTION:
${instruction}

🧠 ΜΝΗΜΗ & ΠΡΟΤΙΜΗΣΕΙΣ ΧΡΗΣΤΗ (Ακολούθησέ τις πιστά):
${config.agentMemory || 'Δεν έχουν οριστεί ειδικές προτιμήσεις.'}

🎯 ΚΑΘΗΚΟΝ:
Απάντησε ΑΥΣΤΗΡΑ με JSON (χωρίς markdown):
{
 "intent": "draft" ή "question",
 "content": "Το περιεχόμενο απάντησης"
}

📝 ΚΑΝΟΝΕΣ:
• Αν intent="draft": Γράψε ΕΤΟΙΜΟ επαγγελματικό email απαντώντας στο thread με τόνο: ${config.tone}.
• ΠΡΟΣΘΕΤΗ ΟΔΗΓΙΑ: ${config.customPrompt || 'Καμία.'}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const parts = [{ text: systemPrompt }];
        if (audioObj?.data) parts.push({ inlineData: { mimeType: audioObj.mimeType, data: audioObj.data } });

        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts }],
                generationConfig: { temperature: 0.3 } //responseMimeType: "application/json",
            })
        });

        const data = await res.json();
        let raw = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        raw = raw.replace(/```json/g, '').replace(/```/g, '').trim();
        const parsed = JSON.parse(raw);

        if (parsed.intent === 'question') {
            document.getElementById('answerText').innerText = parsed.content;
            navigate('view-answer');
        } else {
            document.getElementById('draftTextarea').value = parsed.content;
            navigate('view-draft');
        }
        voiceStatus.innerText = 'Κάντε κλικ για ομιλία';
    } catch (e) {
        alert("Σφάλμα μηχανής AI: " + e.message);
        voiceStatus.innerText = 'Κάντε κλικ για ομιλία';
    }
}

// ----------------------
// CORE INTERACTION EVENT BINDINGS
// ----------------------
function handleQuickAction(actionType) {
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
    generateDraft(`Τροποποίησε το προηγούμενο draft email σύμφωνα με: "${tweak}"\n\nΠΑΛΙΟ EMAIL:\n${current}`, null);
};

// ----------------------
// VOICE CONTROLS
// ----------------------
const voiceBtn = document.getElementById('voiceBtn');
const voiceStatus = document.getElementById('voiceStatus');

voiceBtn.onclick = () => {
    if (isRecording) { stopRecording(); return; }
    startRecording();
};

function startRecording() {
    navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        mediaRecorder = new MediaRecorder(stream, { mimeType: 'audio/webm' });
        audioChunks = [];
        mediaRecorder.ondataavailable = e => { if (e.data.size > 0) audioChunks.push(e.data); };
        
        mediaRecorder.onstop = () => {
            voiceStatus.innerText = '🔄 Επεξεργασία ήχου...';
            const blob = new Blob(audioChunks, { type: 'audio/webm' });
            const reader = new FileReader();
            reader.readAsDataURL(blob);
            reader.onloadend = () => {
                const base64 = reader.result.split(',')[1];
                generateDraft('🎤 Φωνητική εντολή χρήστη', { data: base64, mimeType: 'audio/webm' });
            };
        };

        mediaRecorder.start(1000);
        isRecording = true;
        voiceBtn.className = "w-24 h-24 rounded-full siri-listening flex items-center justify-center text-primary cursor-pointer border border-border";
        voiceBtn.innerHTML = `<i data-lucide="square" class="w-8 h-8 text-white"></i>`;
        lucide.createIcons();
        voiceStatus.innerText = '🔴 Καταγραφή... Μιλήστε τώρα';
    }).catch(() => { voiceStatus.innerText = '❌ Σφάλμα μικροφώνου'; });
}

function stopRecording() {
    if (mediaRecorder && mediaRecorder.state !== 'inactive') {
        mediaRecorder.stop();
        mediaRecorder.stream.getTracks().forEach(track => track.stop());
    }
    isRecording = false;
    voiceBtn.className = "w-24 h-24 rounded-full siri-idle flex items-center justify-center text-primary cursor-pointer border border-border";
    voiceBtn.innerHTML = `<i data-lucide="mic" class="w-8 h-8 opacity-70"></i>`;
    lucide.createIcons();
    voiceStatus.innerText = 'Κάντε κλικ για ομιλία';
}

// ----------------------
// INSERT COMPONENT TO OUTLOOK
// ----------------------
document.getElementById('insertOutlookBtn').onclick = () => {
    const finalTxt = document.getElementById('draftTextarea').value;
    if (!finalTxt.trim()) return;
    
    Office.context.mailbox.item.displayReplyForm(finalTxt, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById('draftTextarea').value = '';
            navigate('view-main');
        } else {
            alert("Σφάλμα αυτόματης επικόλλησης.");
        }
    });
};

document.getElementById('cancelDraftBtn')?.addEventListener('click', () => {
    document.getElementById('draftTextarea').value = '';
    navigate('view-main');
});

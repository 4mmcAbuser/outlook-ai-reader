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
    agentMemory: '' 
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
                setTimeout(() => initOutlookData(), 1000);
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
// UTILITIES & BULLETPROOF JSON PARSER WITH TEXT FALLBACK
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

// Upgraded Bulletproof Parser: Converts raw text responses into synthetic JSON objects dynamically
// Upgraded Bulletproof Parser: Converts raw text responses into synthetic JSON objects dynamically
function safeJsonParse(rawStr) {
    if (!rawStr) throw new Error("Κενή απάντηση από το AI.");
    
    let cleanStr = rawStr.replace(/```json/gi, '').replace(/```/gi, '').trim();
    let start = cleanStr.indexOf('{');
    
    // Αν το μοντέλο επέστρεψε σκέτο κείμενο (reasoning) χωρίς άγκιστρα
    if (start === -1) {
        console.warn("⚠️ Το AI δεν επέστρεψε JSON δομή. Καθαρισμός reasoning...");
        // Αν υπάρχουν bullet points, τα αφαιρούμε για να μείνει μόνο το καθαρό κείμενο
        let cleanText = cleanStr.replace(/^\s*[\*\-\>]\s*/gm, '').trim();
        return {
            summary: cleanText,
            content: cleanText, 
            intent: (cleanText.length > 120 || cleanText.includes("Αγαπητέ") || cleanText.includes("Γεια")) ? "draft" : "question",
            category: "High Priority"
        };
    }
    
    let braceCount = 0;
    let inString = false;
    let escaping = false;
    let endIndex = -1;
    
    for (let i = start; i < cleanStr.length; i++) {
        const char = cleanStr[i];
        
        if (escaping) { escaping = false; continue; }
        if (char === '\\') { escaping = true; continue; }
        if (char === '"') { inString = !inString; continue; }
        
        if (!inString) {
            if (char === '{') braceCount++;
            else if (char === '}') {
                braceCount--;
                if (braceCount === 0) {
                    endIndex = i;
                    const candidate = cleanStr.substring(start, i + 1);
                    try {
                        return JSON.parse(candidate);
                    } catch (e) {
                        // Αν αποτύχει, συνεχίζουμε για να βρούμε το επόμενο κλεισιμο
                    }
                }
            }
        }
    }
    
    const jsonMatch = cleanStr.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
        try {
            return JSON.parse(jsonMatch[0]);
        } catch (e) {
            // Αν σπάσει πάλι, επιστρέφουμε καθαρό κείμενο
        }
    }
    
    // Fallback: καθαρίζουμε bullets και επιστρέφουμε απλό κείμενο
    let fallbackText = cleanStr.replace(/^\s*[\*\-\>]\s*/gm, '').trim();
    return {
        summary: fallbackText,
        content: fallbackText,
        intent: "draft",
        category: "High Priority"
    };
}

// ----------------------
// TEXT TO SPEECH (FOR CLIENT'S HEADPHONES)
// ----------------------
function speakSummary() {
    if (!window.speechSynthesis) {
        document.getElementById('voiceStatus').innerText = "⚠️ Το σύστημα δεν υποστηρίζει Text-to-Speech.";
        return;
    }

    if (window.speechSynthesis.speaking) {
        window.speechSynthesis.cancel();
        document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Ακρόαση`;
        lucide.createIcons();
        return;
    }

    const textToRead = document.getElementById('summaryText').innerText;
    if (textToRead === 'Περιμένω ανάλυση...' || textToRead.startsWith('⚠️') || textToRead.startsWith('⏳') || textToRead.includes('Αδυναμία')) return;

    currentSpeechUtterance = new SpeechSynthesisUtterance(textToRead);
    currentSpeechUtterance.lang = 'el-GR'; 
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

    emailContext.text = '';
    emailContext.fullConversation = [];

    emailContext.meta = {
        senderName: item.sender?.displayName || 'Άγνωστος',
        senderEmail: item.sender?.emailAddress || '',
        subject: item.subject || '(Χωρίς θέμα)',
        receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('el-GR') : ''
    };

    document.getElementById('voiceStatus').innerText = 'Κάντε κλικ για ομιλία'; 
    startLoadingAnim(["📡 Σύνδεση...", "🔍 Ανάλυση Thread...", "🤖 Κατηγοριοποίηση..."]);

    getFullConversationViaREST()
        .then(messages => {
            emailContext.fullConversation = messages;
            const structured = messages.map((m, idx) => `--- ΜΗΝΥΜΑ #${idx + 1} ---\nΑΠΟΣΤΟΛΕΑΣ: ${m.sender}\nΘΕΜΑ: ${m.subject}\nΠΕΡΙΕΧΟΜΕΝΟ:\n${m.body}\n`).join("\n");
            emailContext.text = structured;
            finishLoading();
        })
        .catch(() => {
            fallbackCurrentMail(0);
        });
}

function fallbackCurrentMail(retryCount = 0) {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) {
        stopLoadingAnim();
        emailContext.text = "Σφάλμα ανάγνωσης.";
        finishLoading();
        return;
    }
    
    item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.trim().length > 0) {
            stopLoadingAnim();
            emailContext.text = `--- ΤΡΕΧΟΝ ΜΗΝΥΜΑ ---\nΑΠΟΣΤΟΛΕΑΣ: ${emailContext.meta.senderName}\nΘΕΜΑ: ${emailContext.meta.subject}\n\n${result.value}`;
            finishLoading();
        } else {
            if (retryCount < 3) {
                console.log(`⏳ Άδειο περιεχόμενο. Επαναδοκιμή φόρτωσης... (${retryCount + 1}/3)`);
                setTimeout(() => fallbackCurrentMail(retryCount + 1), 500);
            } else {
                stopLoadingAnim();
                emailContext.text = `--- ΤΡΕΧΟΝ ΜΗΝΥΜΑ ---\nΑΠΟΣΤΟΛΕΑΣ: ${emailContext.meta.senderName}\nΘΕΜΑ: ${emailContext.meta.subject}\n\n[Δεν εντοπίστηκε κείμενο στο email]`;
                finishLoading();
            }
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
// EMAIL AUDIT & SUMMARY ENGINE
// ----------------------
async function generateSummaryAndAudit() {
    if (!config.apiKey) {
        document.getElementById('summaryText').innerText = '⚠️ Εκκρεμεί το API Key στις ρυθμίσεις.';
        return;
    }

    if (!emailContext.text || emailContext.text.trim() === "" || emailContext.text.includes("Σφάλμα ανάγνωσης.")) {
        document.getElementById('summaryText').innerText = '⏳ Αναμονή για φόρτωση email...';
        return;
    }

    const prompt = `Είσαι ένα αυτοματοποιημένο backend API. Η απάντησή σου πρέπει να είναι ΑΠΟΚΛΕΙΣΤΙΚΑ ένα έγκυρο αντικείμενο JSON. Απαγορεύεται ρητά να γράψεις εισαγωγικό κείμενο, markdown, αποσιωπητικά ή επεξηγήσεις.

Ανάλυσε το παρακάτω email thread και συμπλήρωσε τα δεδομένα ακριβώς σε αυτή τη δομή JSON:
{
  "summary": "Μια σύντομη executive σύνοψη (έως 4 γραμμές) στα Ελληνικά.",
  "category": "High Priority"
}

(Για το πεδίο category επίλεξε υποχρεωτικά μία από αυτές τις 4 τιμές: "High Priority", "Internal", "Newsletter", "Spam")

EMAIL THREAD ΓΙΑ ΑΝΑΛΥΣΗ:
${emailContext.text.substring(0, 8000)}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.1 }
            })
        });

        if (!res.ok) {
            const errBody = await res.json().catch(() => ({}));
            throw new Error(`HTTP ${res.status}: ${errBody.error?.message || res.statusText}`);
        }

        const data = await res.json();
        if (data.error) throw new Error(`API Error: ${data.error.message}`);

        let raw = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        
        // Καλούμε τον έξυπνο parser που αντέχει plain κείμενα
        let resObj = safeJsonParse(raw);
        
        if (!resObj.summary) throw new Error("Το JSON δεν περιέχει το πεδίο 'summary'.");
        
        // Καθαρή εμφάνιση στη σύνοψη
        document.getElementById('summaryText').innerText = resObj.summary;
        updateAuditMetrics(resObj.category || 'Spam');
        renderCategoryBadge(resObj.category || 'Spam');

        if (resObj.summary.length > 150) {
            document.getElementById('summaryContent').classList.add('max-h-24');
            document.getElementById('summaryFade')?.classList.remove('hidden');
            document.getElementById('expandSummaryBtn')?.classList.remove('hidden');
        }
    } catch (e) {
        console.error("❌ Summary Generation Error:", e);
        // 🔥 FIX: Η κάρτα της σύνοψης παραμένει καθαρή χωρίς data errors
        document.getElementById('summaryText').innerText = 'Αδυναμία αυτόματης φόρτωσης σύνοψης.';
        // 🔥 FIX: Το τεχνικό σφάλμα απομονώνεται αυστηρά κάτω από το μικρόφωνο
        document.getElementById('voiceStatus').innerText = `⚠️ Σφάλμα Σύνοψης: ${e.message}`;
    }
}

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

document.getElementById('manualSummaryBtn')?.addEventListener('click', () => {
    generateSummaryAndAudit();
});

// ----------------------
// GENERATE DRAFT WITH AGENTIC MEMORY
// ----------------------
// ----------------------
// GENERATE DRAFT WITH AGENTIC MEMORY (OPTIMIZED AGAINST 503 OVERLOAD)
// ----------------------
async function generateDraft(instruction, audioObj) {
    if (!config.apiKey) { navigate('view-settings'); return; }

    if (!emailContext.text || emailContext.text.trim() === "" || emailContext.text.includes("Σφάλμα ανάγνωσης.")) {
        document.getElementById('voiceStatus').innerText = "⏳ Το email φορτώνει ακόμα... Ξαναπροσπαθήστε σε 1 δευτερόλεπτο.";
        return;
    }

    document.getElementById('voiceStatus').innerText = '🤖 Σκέφτομαι...';

    const optimizedContext = emailContext.text.length > 10000
        ? emailContext.text.substring(0, 10000) + "\n\n[...το παλαιότερο ιστορικό περικόπηκε...]"
        : emailContext.text;

    // 🔥 ΠΑΡΑΔΕΙΓΜΑ ΠΡΟΣ ΜΙΜΗΣΗ (FEW-SHOT) ΓΙΑ ΝΑ ΚΑΤΑΛΑΒΕΙ ΟΤΙ ΔΕΝ ΘΕΛΟΥΜΕ REASONING
    const systemPrompt = `Είσαι ένας Executive AI Assistant για επαγγελματική αλληλογραφία.
    
📧 ΙΣΤΟΡΙΚΟ ΣΥΝΟΜΙΛΙΑΣ (EMAIL THREAD):
 ${optimizedContext}

👤 ΟΔΗΓΙΑ ΧΡΗΣΤΗ:
 ${instruction}

🧠 ΜΝΗΜΗ & ΠΡΟΤΙΜΗΣΕΙΣ ΧΡΗΣΤΗ:
 ${config.agentMemory || 'Δεν έχουν οριστεί ειδικές προτιμήσεις.'}

🎯 ΚΑΘΗΚΟΝ:
Επέστρεψε ΑΥΣΤΗΡΑ ένα JSON. 
ΑΠΑΓΟΡΕΥΕΤΑΙ να γράψεις σκέψεις, λογική, ανάλυση, reasoning ή bullet points. Η έξοδος πρέπει να ξεκινάει απευθείας με "{".

Δομή JSON:
{
  "intent": "draft" αν φτιάχνεις email, ή "question" αν ο χρήστης κάνει γενική ερώτηση,
  "content": "Γράψε εδώ απευθείας το πλήρες, ολοκληρωμένο κείμενο της απάντησης ή του email στα Ελληνικά. Το κείμενο πρέπει να είναι έτοιμο προς αποστολή."
}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const parts = [{ text: systemPrompt }];
        if (audioObj?.data) parts.push({ inlineData: { mimeType: audioObj.mimeType, data: audioObj.data } });

        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts }],
                generationConfig: { 
                    temperature: 0.2,
                    responseMimeType: "application/json", // Εξασφαλίζει έγκυρο JSON
                    maxOutputTokens: 2048 // Δίνει χώρο στο AI να γράψει ένα μεγάλο email
                }
            })
        });

        if (!res.ok) {
            const errBody = await res.json().catch(() => ({}));
            throw new Error(`HTTP ${res.status}: ${errBody.error?.message || res.statusText}`);
        }

        const data = await res.json();
        let raw = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        
        const parsed = safeJsonParse(raw);

        // Επιπλέον προστασία: Αν το content περιέχει σκέψεις (π.χ. * Role:...) τις καθαρίζουμε
        let finalContent = parsed.content || "";
        if (finalContent.startsWith("*") || finalContent.startsWith("-")) {
            // Κόβουμε τα bullets μέχρι να φτάσουμε στο πραγματικό κείμενο
            const lines = finalContent.split('\n');
            const cleanLines = lines.filter(line => !line.trim().startsWith("*") && !line.trim().startsWith("-"));
            finalContent = cleanLines.join('\n').trim();
        }

        if (parsed.intent === 'question') {
            document.getElementById('answerText').innerText = finalContent;
            navigate('view-answer');
        } else {
            document.getElementById('draftTextarea').value = finalContent;
            navigate('view-draft');
        }
        document.getElementById('voiceStatus').innerText = 'Κάντε κλικ για ομιλία';
    } catch (e) {
        console.error("🤖 AI Draft Error:", e);
        document.getElementById('voiceStatus').innerText = '❌ Σφάλμα AI: ' + e.message;
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
            console.error(asyncResult.error);
            voiceStatus.innerText = "⚠️ Αποτυχία αυτόματης επικόλλησης.";
        }
    });
};

document.getElementById('cancelDraftBtn')?.addEventListener('click', () => {
    document.getElementById('draftTextarea').value = '';
    navigate('view-main');
});

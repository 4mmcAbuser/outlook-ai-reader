// Initialize Lucide Icons
lucide.createIcons();

// ----------------------
// STATE MANAGEMENT
// ----------------------
let config = {
    apiKey: '',
    model: 'gemini-2.5-flash', // Προεπιλογή σε σταθερό Gemini
    tone: 'Επαγγελματικός, ευγενικός και σοβαρός.',
    autoSum: true,
    customPrompt: '',
    agentMemory: '' 
};

let emailContext = {
    text: '',
    meta: {},
    fullConversation: [],
    myEmail: '',   // Το email του ίδιου του χρήστη
    myName: '',    // Το όνομα του ίδιου του χρήστη
    myDomain: ''   // Το εταιρικό domain του οργανισμού
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
                if (window.speechSynthesis) window.speechSynthesis.cancel(); 
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
// UTILITIES & BULLETPROOF JSON PARSER
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

function safeJsonParse(rawStr) {
    if (!rawStr) throw new Error("Κενή απάντηση από το AI.");
    
    let cleanStr = rawStr.replace(/```json/gi, '').replace(/```/gi, '').trim();
    let start = cleanStr.indexOf('{');
    
    if (start === -1) {
        return {
            summary: cleanStr,
            content: cleanStr, 
            intent: "draft",
            category: "High Priority"
        };
    }
    
    let braceCount = 0;
    let inString = false;
    let escaping = false;
    
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
                    const candidate = cleanStr.substring(start, i + 1);
                    return JSON.parse(candidate);
                }
            }
        }
    }
    
    const jsonMatch = cleanStr.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
        return {
            summary: cleanStr,
            content: cleanStr,
            intent: "draft",
            category: "High Priority"
        };
    }
    return JSON.parse(jsonMatch[0]);
}

// ----------------------
// TEXT TO SPEECH (FOR HEADPHONES)
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

    // 🚀 ΑΝΑΒΑΘΜΙΣΗ 1 & 2: Αυτόματη άντληση στοιχείων χρήστη και domain οργανισμού
    if (Office.context.mailbox.userProfile) {
        emailContext.myEmail = Office.context.mailbox.userProfile.emailAddress || '';
        emailContext.myName = Office.context.mailbox.userProfile.displayName || '';
        if (emailContext.myEmail.includes('@')) {
            emailContext.myDomain = emailContext.myEmail.split('@')[1].toLowerCase();
        }
    }

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
// EMAIL AUDIT & SUMMARY ENGINE ($1,000,000 PROMPT)
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

    // 🔥 ΜΕΓΑΛΗ ΑΝΑΒΑΘΜΙΣΗ: Το AI γνωρίζει την ταυτότητά μας, το domain μας ΚΑΙ τη Μνήμη του Agent για το Audit!
    const prompt = `You are a world-class $1M Corporate Executive Assistant and Chief of Staff. 
Your identity profile:
- Your User's Name: "${emailContext.myName}"
- Your User's Email: "${emailContext.myEmail}"
- Internal Corporate Domain: "${emailContext.myDomain}" (Any email ending with @${emailContext.myDomain} is an INTERNAL team member/colleague. Other domains are external clients or spam).

USER STRATEGIC MEMORY & PREFERENCES (Apply strictly to categorization and summary logic):
${config.agentMemory || 'No custom preferences set.'}

TASK:
Analyze the provided email thread. Determine if the latest communication comes from an internal colleague or an external stakeholder. Output a strictly valid JSON object ONLY. No markdown, no conversations.

JSON Structure:
{
  "summary": "Μια κορυφαία, εξαιρετικά περιεκτική executive σύνοψη (έως 4 γραμμές) στα Ελληνικά. Πρέπει να αναφέρει ποιος έστειλε το τελευταίο μήνυμα, τι ακριβώς ζητάει, και αν εκκρεμεί δική μας ενέργεια.",
  "category": "High Priority"
}

Categorization Rules (Choose EXACTLY one of these four tokens):
- "High Priority": Crucial external client requests, financial decisions, urgent deliverables, or alerts matching the User Strategic Memory profile.
- "Internal": Messages sent from colleagues inside the @${emailContext.myDomain} domain that don't require emergency external client attention.
- "Newsletter": News, updates, reports, circulars, or automated tracking alerts.
- "Spam": Unimportant, cold sales pitches, or clutter.

EMAIL THREAD TO PROCESS:
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
        let raw = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        
        let resObj = safeJsonParse(raw);
        
        document.getElementById('summaryText').innerText = resObj.summary;
        updateAuditMetrics(resObj.category || 'Spam');
        renderCategoryBadge(resObj.category || 'Spam');

        if (resObj.summary.length > 150) {
            document.getElementById('summaryContent').classList.add('max-h-24');
            document.getElementById('summaryFade')?.classList.remove('hidden');
            document.getElementById('expandSummaryBtn')?.classList.remove('hidden');
        }
    } catch (e) {
        console.error("❌ Summary Error:", e);
        document.getElementById('summaryText').innerText = 'Αδυναμία αυτόματης φόρτωσης σύνοψης.';
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
// GENERATE DRAFT WITH AGENTIC MEMORY ($1,000,000 PROMPT)
// ----------------------
async function generateDraft(instruction, audioObj) {
    if (!config.apiKey) { navigate('view-settings'); return; }

    if (!emailContext.text || emailContext.text.trim() === "" || emailContext.text.includes("Σφάλμα ανάγνωσης.")) {
        document.getElementById('voiceStatus').innerText = "⏳ Το email φορτώνει ακόμα... Ξαναπροσπαθήστε σε 1 δευτερόλεπτο.";
        return;
    }

    document.getElementById('voiceStatus').innerText = '🤖 Σκέφτομαι...';

    const optimizedContext = emailContext.text.length > 10000
        ? emailContext.text.substring(0, 10000)
        : emailContext.text;

    // 🔥 ΥΠΕΡ-ΠΡΟΜΠΤ 1.000.000$: Καθαρό Command Structure, απαγόρευση Chain-of-Thought (Checklists), πλήρης επίγνωση ταυτότητας
    const systemPrompt = `You are an elite, multi-million dollar Executive AI Agent and senior corporate diplomat crafting corporate correspondence.
Your User Profile Context (You are writing on behalf of this person):
- User's Name: "${emailContext.myName}"
- User's Email Address: "${emailContext.myEmail}"
- Corporate Organization Domain: "@${emailContext.myDomain}" (Colleagues match this domain, external clients/leads use different domains).

CRITICAL ASSIGNMENT:
Generate a single valid JSON object and absolutely NOTHING else. No markdown wrappers, no introductory comments, no thought processes, no internal checklists.

JSON Schema Requirement:
{"intent": "draft", "content": "Your flawless, ready-to-insert Greek email body text goes here"}

Strict Execution Protocols:
1. Write the final email draft directly in the "content" field.
2. The response must be completely written in fluid, natural Greek, aligning with this specific tone directive: "${config.tone}".
3. PERMANENT USER MEMORY RULES (Prioritize these above all text generation guidelines):
${config.agentMemory || 'No strategic user memory rules have been provided yet.'}
4. Additional Custom Injection Rules: ${config.customPrompt || 'None.'}
5. Differentiate identity: Analyze the thread. Any message sent from "${emailContext.myEmail}" was sent by YOU. Do not treat your own past words as the client's words.
6. Do not include subject lines, placeholders, empty bracket templates, or ellipses (...). Write a polished, premium, ready-to-send message.
7. If the context thread appears blank, act fully autonomously to craft a brilliant independent email meeting the user command.

EMAIL THREAD CONTEXT:
${optimizedContext}

USER EXECUTIVE COMMAND:
${instruction}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const parts = [{ text: systemPrompt }];
        if (audioObj?.data) parts.push({ inlineData: { mimeType: audioObj.mimeType, data: audioObj.data } });

        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts }],
                generationConfig: { temperature: 0.2 }
            })
        });

        if (!res.ok) {
            const errBody = await res.json().catch(() => ({}));
            throw new Error(`HTTP ${res.status}: ${errBody.error?.message || res.statusText}`);
        }

        const data = await res.json();
        let raw = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        
        const parsed = safeJsonParse(raw);

        if (!parsed.content || parsed.content.trim() === "..." || parsed.content.trim().length < 5) {
            console.warn("⚠️ Fallback activated.");
            parsed.content = `Αγαπητέ συνεργάτη,\n\nΣε συνέχεια του μηνύματός σας, θα ήθελα να σας ενημερώσω ότι αποδέχομαι την πρόταση / προσφορά. Θα επανέλθω σύντομα με τις σχετικές λεπτομέρειες.\n\nΜε εκτίμηση`;
        }

        if (parsed.intent === 'question') {
            document.getElementById('answerText').innerText = parsed.content;
            navigate('view-answer');
        } else {
            document.getElementById('draftTextarea').value = parsed.content;
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

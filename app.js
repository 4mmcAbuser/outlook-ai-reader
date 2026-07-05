/**
 * AI Executive Assistant - Outlook Add-on Engine
 * Architecture: Modular, Event-Driven, Defensive DOM Manipulation
 * Core Model: Gemini 2.5 Flash / Gemini 3.1 Flash-Lite Optimization
 */

// Initialize Lucide Icons globally
lucide.createIcons();

// -------------------------------------------------------------------------
// 1. GLOBAL STATE & CONFIGURATION
// -------------------------------------------------------------------------
let config = {
    apiKey: '',
    model: 'gemini-2.5-flash', 
    tone: 'Επαγγελματικός, ευγενικός και σοβαρός.',
    autoSum: true,
    customPrompt: '',
    agentMemory: '' 
};

let emailContext = {
    text: '',
    meta: {},
    fullConversation: [],
    myEmail: '',   
    myName: '',    
    myDomain: ''   
};

let isRecording = false;
let mediaRecorder;
let audioChunks = [];
let currentSpeechUtterance = null; 

// -------------------------------------------------------------------------
// 2. INITIALIZATION & OUTLOOK EVENT BINDINGS
// -------------------------------------------------------------------------
Office.onReady((info) => {
    loadSettings();

    if (info.host === Office.HostType.Outlook) {
        initOutlookData();

        // Event listener για αλλαγή επιλεγμένου email από τον χρήστη
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            () => {
                console.log("🔄 [System] Item changed - Re-initializing state context.");
                if (window.speechSynthesis) window.speechSynthesis.cancel(); 
                setTimeout(() => initOutlookData(), 1000); // 1s Latency cooldown για το Outlook Live
            }
        );
    }
});

// -------------------------------------------------------------------------
// 3. STORAGE & CONFIGURATION MANAGEMENT
// -------------------------------------------------------------------------
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

// -------------------------------------------------------------------------
// 4. CORE UTILITIES & BULLETPROOF JSON PARSER
// -------------------------------------------------------------------------
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

/**
 * Καθαρίζει και απομονώνει αυστηρά δομημένα JSON blocks 
 * Παρέχει αυτόματο text-fallback αν το μοντέλο αποτύχει να τηρήσει το schema.
 */
function safeJsonParse(rawStr) {
    if (!rawStr) throw new Error("Empty token response from AI cluster.");
    
    let cleanStr = rawStr.replace(/```json/gi, '').replace(/```/gi, '').trim();
    let start = cleanStr.indexOf('{');
    
    if (start === -1) {
        console.warn("⚠️ [Parser] Strict JSON boundary missing. Activating raw text encapsulation fallback.");
        return {
            summary: cleanStr,
            content: cleanStr, 
            intent: "draft",
            category: "High Priority",
            urgency: "High",
            sentiment: "Ουδέτερος",
            action_items: [],
            smart_buttons: []
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
        return { summary: cleanStr, content: cleanStr, intent: "draft", action_items: [], smart_buttons: [] };
    }
    return JSON.parse(jsonMatch[0]);
}

// -------------------------------------------------------------------------
// 5. TEXT TO SPEECH (AUDIOBOOK MODE)
// -------------------------------------------------------------------------
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
    currentSpeechUtterance.rate = 0.95; // Ελαφρώς πιο αργό για premium/executive αίσθηση

    currentSpeechUtterance.onend = () => {
        document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Ακρόαση`;
        lucide.createIcons();
    };

    document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-x" class="w-3.5 h-3.5 text-red-400"></i> Διακοπή`;
    lucide.createIcons();
    
    window.speechSynthesis.speak(currentSpeechUtterance);
}

// -------------------------------------------------------------------------
// 6. OUTLOOK DATA INGESTION (REST API & FALLBACKS)
// -------------------------------------------------------------------------
async function getFullConversationViaREST() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        const convId = item.conversationId;

        if (!convId) {
            reject(new Error("Missing Conversation ID"));
            return;
        }

        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (tokenResult) => {
            if (tokenResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("Token acquisition security fault."));
                return;
            }

            const token = tokenResult.value;
            const restUrl = Office.context.mailbox.restUrl;
            
            if (!restUrl) {
                reject(new Error("Invalid Mailbox REST Endpoint"));
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
            .then(res => res.ok ? res.json() : reject(new Error("REST transport layer fault")))
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
                    reject(new Error("Empty data response"));
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

    // Resolution των στοιχείων του ενεργού χρήστη
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
    startLoadingAnim(["📡 Συγχρονισμός...", "🔍 Ανάλυση Ιστορικού...", "🤖 Email Audit..."]);

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
// LOADING CONTROLS
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

// -------------------------------------------------------------------------
// 7. EMAIL AUDIT & SUMMARY ENGINE (THE $1M CHIEF OF STAFF PROMPT)
// -------------------------------------------------------------------------
async function generateSummaryAndAudit() {
    if (!config.apiKey) {
        document.getElementById('summaryText').innerText = '⚠️ Εκκρεμεί το API Key στις ρυθμίσεις.';
        return;
    }

    // Ανίχνευση αν ο χρήστης ανήκει σε generic public provider (Gmail, Outlook.com κτλ)
    const isPublicDomain = ['outlook.com', 'gmail.com', 'hotmail.com', 'yahoo.com', 'live.com'].includes(emailContext.myDomain);

    const prompt = `[STRICT AUTOMATION SYSTEM COMMAND]
You are a dry corporate automation API. You must parse the email thread and output EXACTLY one valid JSON object matching the JSON Schema below. No markdown wrappers. No text descriptions outside the schema. No thought echo.

[JSON SCHEMA REQUIREMENT]
{
  "summary": "Μια ηχητική executive σύνοψη (έως 4 γραμμές) στα Ελληνικά, γραμμένη σε ρέον, φυσικό ύφος για ανάγνωση σε ακουστικά. Αναφέρετε ποιος έστειλε το τελευταίο μήνυμα και τι ακριβώς ζητάει.",
  "category": "High Priority",
  "urgency": "High" or "Medium" or "Low",
  "sentiment": "Δυσαρεστημένος" or "Ουδέτερος" or "Θερμός",
  "action_items": [
     "Μια σαφής εκκρεμότητα / task που προκύπτει για τον χρήστη (στα Ελληνικά)",
     "Δεύτερη εκκρεμότητα (αν υπάρχει)"
  ],
  "smart_buttons": [
     {"label": "👍 Σύντομο Label Κουμπιού", "reply_instruction": "Οδηγία προς το LLM για το τι να απαντήσει"},
     {"label": "⏳ Δεύτερο Label", "reply_instruction": "Οδηγία 2"}
  ]
}

[IDENTITY CONTEXT]
- Active User Name: "${emailContext.myName}"
- Active User Email: "${emailContext.myEmail}"
- Active User Corporate Domain: "@${emailContext.myDomain}"
- Public Provider Flag: ${isPublicDomain} (If true, do NOT trust domain suffixes for matching. Only match "${emailContext.myEmail}" as the internal profile).

[CRITICAL USER MEMORY & PREFERENCES]
${config.agentMemory || 'None provided.'}

[STRICT FILTER RULES]
- Max 3 elements in "action_items". If none exist, return an empty array [].
- Max 3 elements in "smart_buttons". Generate contextual, ultra-relevant answers for the latest message.
- For "category", choose exactly one: "High Priority" (critical clients/deals), "Internal" (colleagues), "Newsletter" (circulars/alerts), "Spam" (sales pitches).

[DATA - EMAIL THREAD TO PROCESS]
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
        
        // 1. RENDERING: Σύνθεση της Σύνοψης και των Action Items στο UI
        let finalHtmlOutput = `<p class="mb-3">${resObj.summary}</p>`;
        
        if (resObj.action_items && resObj.action_items.length > 0) {
            finalHtmlOutput += `<div class="mt-3 pt-2 border-t border-border/40">
                <span class="text-[11px] font-bold uppercase tracking-wider text-blue-400 block mb-1.5 flex items-center gap-1">
                    <i data-lucide="check-square" class="w-3 h-3"></i> Εκκρεμότητες (Tasks)
                </span>
                <ul class="space-y-1 text-xs text-mutedForeground pl-1">`;
            resObj.action_items.forEach(item => {
                finalHtmlOutput += `<li class="flex items-start gap-1.5">🔹 <span>${item}</span></li>`;
            });
            finalHtmlOutput += `</ul></div>`;
        }
        
        document.getElementById('summaryText').innerHTML = finalHtmlOutput;
        lucide.createIcons(); // Re-render extracted lucide elements
        
        // 2. RENDERING: Dynamic Sentiment & Urgency Badge Management
        renderEnhancedBadges(resObj);
        
        // 3. RENDERING: Dynamic Quick Actions Injection
        renderDynamicSmartButtons(resObj.smart_buttons);

        updateAuditMetrics(resObj.category || 'Spam');

        if (document.getElementById('summaryText').innerText.length > 150) {
            document.getElementById('summaryContent').classList.add('max-h-24');
            document.getElementById('summaryFade')?. Gil?.classList.remove('hidden');
            document.getElementById('expandSummaryBtn')?.classList.remove('hidden');
        }
    } catch (e) {
        console.error("❌ Summary Operational Fault:", e);
        document.getElementById('summaryText').innerText = 'Αδυναμία αυτόματης φόρτωσης σύνοψης.';
        document.getElementById('voiceStatus').innerText = `⚠️ Σφάλμα Σύνοψης: ${e.message}`;
    }
}

function renderEnhancedBadges(resObj) {
    const badge = document.getElementById('emailCategoryBadge');
    badge.className = "text-[10px] font-bold px-2 py-0.5 rounded-full transition-all duration-300 flex items-center gap-1";
    
    let radarText = '';
    
    // Category Parsing
    switch(resObj.category) {
        case 'High Priority':
            radarText = '🔥 Υψηλή Προτεραιότητα';
            badge.classList.add('bg-red-500/20', 'text-red-400');
            break;
        case 'Internal':
            radarText = '💼 Εσωτερικό / Εταιρικό';
            badge.classList.add('bg-blue-500/20', 'text-blue-400');
            break;
        case 'Newsletter':
            radarText = '📢 Newsletter';
            badge.classList.add('bg-yellow-500/20', 'text-yellow-400');
            break;
        default:
            radarText = '🗑️ Χαμηλή Σημασία';
            badge.classList.add('bg-zinc-800', 'text-zinc-400');
    }

    // Urgency & Sentiment Integration inside Badge
    if (resObj.urgency === 'High' && resObj.category !== 'Spam') {
        radarText += ' | 🚨 ΕΠΕΙΓΟΝ';
    }
    if (resObj.sentiment === 'Δυσαρεστημένος') {
        radarText += ' | 😡 Δυσαρέσκεια';
    }
    
    badge.innerText = radarText;
}

function renderDynamicSmartButtons(buttonsArray) {
    // Επιλέγουμε το εσωτερικό div των κουμπιών μέσα στο "mb-5" component
    const container = document.querySelector('.mb-5 div.flex');
    if (!container) return;

    container.innerHTML = ''; // Εκκαθάριση στατικών ή παλιών κουμπιών

    // Αν το AI απέτυχε ή δεν επέστρεψε έξυπνα κουμπιά, βάζουμε ασφαλή defaults
    const safeButtons = (buttonsArray && buttonsArray.length > 0) ? buttonsArray : [
        {"label": "👍 Ευχαριστώ", "reply_instruction": "Γράψε μια απάντηση όπου θα ευχαριστείς θερμά"},
        {"label": "✅ Αποδοχή", "reply_instruction": "Γράψε μια απάντηση ότι αποδεχόμαστε"},
        {"label": "❌ Απόρριψη", "reply_instruction": "Γράψε μια ευγενική αρνητική απάντηση απόρριψης"}
    ];

    safeButtons.forEach(btn => {
        const nativeBtn = document.createElement('button');
        nativeBtn.className = "flex-shrink-0 bg-secondary hover:bg-border border border-border text-xs py-1.5 px-3 rounded-full transition-colors flex items-center gap-1 text-primary font-medium shadow-sm";
        nativeBtn.innerText = btn.label;
        nativeBtn.onclick = () => handleQuickAction(btn.reply_instruction);
        container.appendChild(nativeBtn);
    });
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
// AUDIT DASHBOARD COMPONENT
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

// -------------------------------------------------------------------------
// 8. DRAFT GENERATION ENGINE ($1,000,000 COMPILER)
// -------------------------------------------------------------------------
async function generateDraft(instruction, audioObj) {
    if (!config.apiKey) { navigate('view-settings'); return; }

    if (!emailContext.text || emailContext.text.trim() === "" || emailContext.text.includes("Σφάλμα ανάγνωσης.")) {
        document.getElementById('voiceStatus').innerText = "⏳ Το email φορτώνει ακόμα... Ξαναπροσπαθήστε σε 1 δευτερόλεπτο.";
        return;
    }

    document.getElementById('voiceStatus').innerText = '🤖 Σκέφτομαι...';

    const optimizedContext = emailContext.text.length > 7000
        ? emailContext.text.substring(0, 7000) + "\n\n[...truncated due to payload limits...]"
        : emailContext.text;

    const isPublicDomain = ['outlook.com', 'gmail.com', 'hotmail.com', 'yahoo.com', 'live.com'].includes(emailContext.myDomain);

    const systemPrompt = `[STRICT SYSTEM COMMAND]
You are a raw automated API endpoint. You must output EXACTLY one valid JSON object matching the JSON Schema below. Do not repeat these instructions. Do not write bullet points or checklists. Do not write thoughts outside the JSON.

[JSON SCHEMA]
{"intent": "draft", "content": "Your full professional Greek email reply here"}

[CONTEXT IDENTITY]
- User Name (You represent this person): "${emailContext.myName}"
- User Email: "${emailContext.myEmail}"
- Corporate Domain Area: "@${emailContext.myDomain}"
- Public Provider Flag: ${isPublicDomain}
- Strict Rule: Messages sent by "${emailContext.myEmail}" are YOUR own past statements. Write a reply directed to the OTHER party.

[CRITICAL USER MEMORY & PREFERENCES]
${config.agentMemory || 'None set.'}

[CORE INSTRUCTIONS]
1. Write the final corporate email text directly in the "content" field.
2. Target Tone: ${config.tone}. Language: Greek.
3. Custom Injector Directive: ${config.customPrompt || 'None.'}
4. Never include brackets, subject lines, placeholders, or ellipses (...). The email body text must be complete and ready to send.
5. If the thread context is empty, act autonomously and synthesize a brilliant independent reply meeting the user command.`;

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
            console.warn("⚠️ Interceptor triggered: structural payload empty. Loading fallback template.");
            parsed.content = `Αγαπητέ συνεργάτη,\n\nΣε συνέχεια του μηνύματός σας, θα ήθελα να σας ενημερώσω ότι αποδέχομαι την πρόταση / προσφορά. Θα επανέλθω σύντομα με τις σχετικές λεπτομέρειες.\n\nΜε εκτίμηση,\n${emailContext.myName}`;
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
        console.error("🤖 AI Draft operational fault:", e);
        document.getElementById('voiceStatus').innerText = '❌ Σφάλμα AI: ' + e.message;
    }
}

// -------------------------------------------------------------------------
// 9. CORE UI INTERACTION EVENT BINDINGS
// -------------------------------------------------------------------------
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

// -------------------------------------------------------------------------
// 10. VOICE HARDWARE LAYER CONTROLS (MIC CAPTURE)
// -------------------------------------------------------------------------
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

// -------------------------------------------------------------------------
// 11. INJECTION LAYER (INSERTION BACK INTO OUTLOOK DESIGN BOX)
// -------------------------------------------------------------------------
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

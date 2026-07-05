/**
 * AI Executive Assistant - Outlook Add-on Engine
 * Performance Profile: Gemini 3.1 Flash-Lite / 2.5 Flash Optimized
 * Design: Dual-Row Contextual Buttons & Native Outlook Calendar Integration
 */

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

        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            () => {
                console.log("🔄 [System] Item changed - Re-initializing state context.");
                if (window.speechSynthesis) window.speechSynthesis.cancel(); 
                setTimeout(() => initOutlookData(), 1000); 
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

function safeJsonParse(rawStr) {
    if (!rawStr) throw new Error("Empty token response from AI cluster.");
    
    let cleanStr = rawStr.replace(/```json/gi, '').replace(/```/gi, '').trim();
    let start = cleanStr.indexOf('{');
    
    if (start === -1) {
        return {
            summary: cleanStr,
            content: cleanStr, 
            intent: "draft",
            category: "High Priority",
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
        return { summary: cleanStr, content: cleanStr, intent: "draft", smart_buttons: [] };
    }
    return JSON.parse(jsonMatch[0]);
}

// ----------------------
// TEXT TO SPEECH
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
    currentSpeechUtterance.rate = 0.95; 

    currentSpeechUtterance.onend = () => {
        document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Ακρόαση`;
        lucide.createIcons();
    };

    document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-x" class="w-3.5 h-3.5 text-red-400"></i> Διακοπή`;
    lucide.createIcons();
    
    window.speechSynthesis.speak(currentSpeechUtterance);
}

// ----------------------
// OUTLOOK API CORE COUPLING
// ----------------------
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
// 3. EMAIL AUDIT & SUMMARY ENGINE (DASHBOARD LOGIC)
// -------------------------------------------------------------------------
async function generateSummaryAndAudit() {
    if (!config.apiKey) {
        document.getElementById('summaryText').innerText = '⚠️ Εκκρεμεί το API Key στις ρυθμίσεις.';
        return;
    }

    const isPublicDomain = ['outlook.com', 'gmail.com', 'hotmail.com', 'yahoo.com', 'live.com'].includes(emailContext.myDomain);

    const prompt = `[STRICT AUTOMATION SYSTEM COMMAND]
You are a dry corporate automation API. You must parse the email thread and output EXACTLY one valid JSON object matching the JSON Schema below. No markdown, no commentary.

[JSON SCHEMA REQUIREMENT]
{
  "summary": "Μια κορυφαία executive σύνοψη (έως 4 γραμμές) στα Ελληνικά, γραμμένη σε φυσικό ύφος για ανάγνωση σε ακουστικά. Αναφέρετε ποιος έστειλε το τελευταίο μήνυμα και τι ακριβώς ζητάει.",
  "category": "High Priority",
  "urgency": "High" or "Medium" or "Low",
  "sentiment": "Δυσαρεστημένος" or "Ουδέτερος" or "Θερμός",
  "detected_meeting_time": "Αν βρεθεί συγκεκριμένη ώρα/ημερομηνία για ραντεβού (π.χ. 'Τρίτη στις 11:00') γράψτην εδώ, αλλιώς κενό ''",
  "smart_buttons": [
     {"label": "📅 Επιβεβαίωση 11:00", "reply_instruction": "Γράψε απάντηση επιβεβαίωσης: Αποδεχόμαστε το ραντεβού για την Τρίτη στις 11:00 ακριβώς. Μην γράψεις γενικόλογα κείμενα."},
     {"label": "⏳ Ζήτα Αλλαγή Ώρας", "reply_instruction": "Γράψε απάντηση ότι η ώρα 11:00 δεν βολεύει και πρότεινε εναλλακτική."}
  ]
}

[IDENTITY CONTEXT]
- Active User Name: "${emailContext.myName}"
- Active User Email: "${emailContext.myEmail}"
- Active User Corporate Domain: "@${emailContext.myDomain}"
- Public Provider Flag: ${isPublicDomain}

[CRITICAL USER MEMORY & FILTERS]
${config.agentMemory || 'None provided.'}

[STRICT FILTER RULES]
- Max 3 custom contextual elements in "smart_buttons". Generate answers targeting the absolute immediate next action of the latest message.

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
        
        // Rendering Summary
        document.getElementById('summaryText').innerHTML = `<p>${resObj.summary}</p>`;
        
        // Rendering Radar Badges
        renderEnhancedBadges(resObj);
        
        // Rendering 2nd Dynamic Button Row + Calendar suggestions
        renderDynamicSmartButtons(resObj.smart_buttons, resObj.detected_meeting_time);

        updateAuditMetrics(resObj.category || 'Spam');

        if (document.getElementById('summaryText').innerText.length > 150) {
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

function renderEnhancedBadges(resObj) {
    const badge = document.getElementById('emailCategoryBadge');
    badge.className = "text-[10px] font-bold px-2 py-0.5 rounded-full transition-all duration-300 flex items-center gap-1";
    let radarText = '';
    
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

    if (resObj.urgency === 'High' && resObj.category !== 'Spam') radarText += ' | 🚨 ΕΠΕΙΓΟΝ';
    if (resObj.sentiment === 'Δυσαρεστημένος') radarText += ' | 😡 Δυσαρέσκεια';
    
    badge.innerText = radarText;
}

// 🚀 FIX: Σχεδιασμός της 2ης σειράς κουμπιών κάτω από τα στατικά με ακριβώς το ίδιο design
function renderDynamicSmartButtons(buttonsArray, detectedTime) {
    const container = document.getElementById('dynamicActionsContainer');
    if (!container) return;

    container.innerHTML = ''; // Καθαρισμός προηγούμενης δυναμικής σειράς
    container.classList.remove('hidden');

    const safeButtons = (buttonsArray && buttonsArray.length > 0) ? buttonsArray : [];

    // 1. Εισαγωγή των έξυπνων AI απαντήσεων
    safeButtons.forEach(btn => {
        const nativeBtn = document.createElement('button');
        // Ακριβές αντίγραφο του design των από πάνω κουμπιών
        nativeBtn.className = "flex-shrink-0 bg-secondary hover:bg-border border border-border text-xs py-1.5 px-3 rounded-full transition-colors text-primary font-medium shadow-sm flex items-center gap-1";
        nativeBtn.innerText = "✨ " + btn.label;
        nativeBtn.onclick = () => handleQuickAction(btn.reply_instruction);
        container.appendChild(nativeBtn);
    });

    // 2. 📅 NATIVE OUTLOOK CALENDAR INTEGRATION: 
    // Αν ανιχνευθεί ώρα, προσθέτουμε ένα ειδικό πράσινο κουμπί άμεσου προγραμματισμού ημερολογίου!
    if (detectedTime && detectedTime.trim() !== '') {
        const calBtn = document.createElement('button');
        calBtn.className = "flex-shrink-0 bg-green-600/20 hover:bg-green-600/40 border border-green-500/40 text-xs py-1.5 px-3 rounded-full transition-colors text-green-400 font-bold shadow-sm flex items-center gap-1 animate-pulse";
        calBtn.innerHTML = `📅 + Ημερολόγιο (${detectedTime})`;
        calBtn.onclick = () => {
            Office.context.mailbox.item.displayNewAppointmentForm({
                subject: `Συνάντηση: ${emailContext.meta.subject}`,
                body: `Προγραμματισμός βάσει του email thread.\n\n${emailContext.text.substring(0, 500)}`,
                notes: "Δημιουργήθηκε αυτόματα από το AI Assistant."
            });
        };
        container.appendChild(calBtn);
    }

    // Αν η σειρά κατέληξε εντελώς άδεια, την κρύβουμε για να μην πιάνει χώρο
    if (container.children.length === 0) {
        container.classList.add('hidden');
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
// 4. DRAFT GENERATION ENGINE (HYPER-STRICT INSTRUCTION FORCE)
// -------------------------------------------------------------------------
async function generateDraft(instruction, audioObj) {
    if (!config.apiKey) { navigate('view-settings'); return; }

    if (!emailContext.text || emailContext.text.trim() === "" || emailContext.text.includes("Σφάλμα ανάγνωσης.")) {
        document.getElementById('voiceStatus').innerText = "⏳ Το email φορτώνει ακόμα... Ξαναπροσπαθήστε.";
        return;
    }

    document.getElementById('voiceStatus').innerText = '🤖 Σκέφτομαι...';

    const optimizedContext = emailContext.text.length > 7000
        ? emailContext.text.substring(0, 7000)
        : emailContext.text;

    // 🔥 FIX: Απόλυτα επιθετικό και αυστηρό prompt. 
    // Καταργεί τις γενικούρες και αναγκάζει το AI να απαντήσει ΣΥΓΚΕΚΡΙΜΕΝΑ στην εντολή του κουμπιού.
    const systemPrompt = `[CRITICAL EXECUTION ORDER - DO NOT HALLUCINATE]
You are a software API backend component that outputs EXACTLY one raw valid JSON object. 
You are prohibited from writing introductory comments, notes, thoughts, or markdown syntax.

[JSON SCHEMA REQUIREMENT]
{"intent": "draft", "content": "Your actual final response email body text goes here"}

[THE ABSOLUTE LAW FOR CONTENT GENERATION]
Your primary directive is to directly execute the [TARGET COMMAND] below. 
- If the [TARGET COMMAND] asks to confirm a specific time (e.g., 'Επιβεβαίωση 11:00'), your generated email text MUST explicitly accept and confirm that exact time (11:00). 
- DO NOT write generic holding responses like "I am looking into it" or "I will reply soon" if the user has already chosen a clear action button.
- Be concise, direct, and elite. Write the email body exactly as the user needs to send it.

[CONTEXT DATA]
- User Name: "${emailContext.myName}"
- User Email: "${emailContext.myEmail}"
- Language: Greek.
- Required Tone Style: ${config.tone}
- User Strategic Memory Rules: ${config.agentMemory || 'None.'}
- Additional Injector Rule: ${config.customPrompt || 'None.'}

[EMAIL THREAD CONTEXT]
${optimizedContext}

[TARGET COMMAND (YOU MUST EXECUTE THIS SPECIFIC ACTION)]
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
                generationConfig: { 
                    temperature: 0.15 // Πολύ χαμηλό για να μην ξεφεύγει από την εντολή του κουμπιού
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

        if (!parsed.content || parsed.content.trim() === "..." || parsed.content.trim().length < 5) {
            parsed.content = `Αγαπητέ συνεργάτη,\n\nΕπιβεβαιώνω τη συνάντησή μας για την προγραμματισμένη ώρα. Θα είμαι διαθέσιμος.\n\nΜε εκτίμηση,\n${emailContext.myName}`;
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
// EVENT BINDINGS
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
// VOICE LAYER
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
// OUTLOOK INJECTION
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

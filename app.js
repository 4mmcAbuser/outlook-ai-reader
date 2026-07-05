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
    tone: '��������������, ��������� ��� �������.',
    autoSum: true,
    autoRead: false,
    ttsRate: '1.0',
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
                console.log("?? [System] Item changed - Re-initializing state context.");
                if (window.speechSynthesis) window.speechSynthesis.cancel(); 
                setTimeout(() => initOutlookData(), 1000); 
            }
        );
    }
});

// ----------------------
// SETTINGS MANAGEMENT
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
    document.getElementById('setAutoRead').checked = !!config.autoRead;
    document.getElementById('setTtsRate').value = config.ttsRate || '1.0';
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
    config.autoRead = document.getElementById('setAutoRead').checked;
    config.ttsRate = document.getElementById('setTtsRate').value;
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
        .replace(/<style[^>]*>[\\s\\S]*?<\\/style>/gi, '')
        .replace(/<script[^>]*>[\\s\\S]*?<\\/script>/gi, '')
        .replace(/<br\\s*\\/?>/gi, '\\n')
        .replace(/<\\/p>/gi, '\\n\\n')
        .replace(/<\\/div>/gi, '\\n')
        .replace(/<div[^>]*>/gi, '')
        .replace(/<[^>]+>/g, '')
        .replace(/\\n{3,}/g, '\\n\\n')
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
        if (char === '\\\\') { escaping = true; continue; }
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
    
    const jsonMatch = cleanStr.match(/\\{[\\s\\S]*\\}/);
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
        document.getElementById('voiceStatus').innerText = "?? �� �穫��� ��� �������坜� Text-to-Speech.";
        return;
    }

    if (window.speechSynthesis.speaking) {
        window.speechSynthesis.cancel();
        document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-2" class="w-3.5 h-3.5"></i> ���昩�`;
        lucide.createIcons();
        return;
    }

    const textToRead = document.getElementById('summaryText').innerText;
    if (textToRead === '�������� ��᢬��...' || textToRead.startsWith('??') || textToRead.startsWith('?') || textToRead.includes('��������')) return;

    currentSpeechUtterance = new SpeechSynthesisUtterance(textToRead);
    currentSpeechUtterance.lang = 'el-GR'; 
    currentSpeechUtterance.rate = parseFloat(config.ttsRate) || 1.0; 

    currentSpeechUtterance.onend = () => {
        document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-2" class="w-3.5 h-3.5"></i> ���昩�`;
        lucide.createIcons();
    };

    document.getElementById('ttsBtn').innerHTML = `<i data-lucide="volume-x" class="w-3.5 h-3.5 text-red-400"></i> �������`;
    lucide.createIcons();
    
    window.speechSynthesis.speak(currentSpeechUtterance);
}

// ----------------------
// OUTLOOK API LAYER
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
                            let senderName = m.Sender?.EmailAddress?.Name || "ꚤ੫��";
                            let senderEmail = m.Sender?.EmailAddress?.Address || "";
                            const bodyContent = m.Body.ContentType === 'HTML' ? cleanHtmlToText(m.Body.Content) : m.Body.Content;
                            
                            return {
                                sender: `${senderName} <${senderEmail}>`,
                                subject: m.Subject || "(����� �⣘)",
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

    // ���������� ��� Reset ��� Summary Box ���� ��� ������ email
    const summaryContent = document.getElementById('summaryContent');
    const expandSummaryBtn = document.getElementById('expandSummaryBtn');
    const summaryFade = document.getElementById('summaryFade');
    if (summaryContent) summaryContent.classList.remove('max-h-24');
    if (expandSummaryBtn) expandSummaryBtn.classList.add('hidden');
    if (summaryFade) summaryFade.classList.add('hidden');

    if (Office.context.mailbox.userProfile) {
        emailContext.myEmail = Office.context.mailbox.userProfile.emailAddress || '';
        emailContext.myName = Office.context.mailbox.userProfile.displayName || '';
        if (emailContext.myEmail.includes('@')) {
            emailContext.myDomain = emailContext.myEmail.split('@')[1].toLowerCase();
        }
    }

    emailContext.meta = {
        senderName: item.sender?.displayName || 'ꚤ੫��',
        senderEmail: item.sender?.emailAddress || '',
        subject: item.subject || '(����� �⣘)',
        receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('el-GR') : ''
    };

    document.getElementById('voiceStatus').innerText = '�ᤫ� ���� ��� ������'; 
    startLoadingAnim(["?? ������������...", "?? ��᢬�� ���������...", "?? Email Audit..."]);

    getFullConversationViaREST()
        .then(messages => {
            emailContext.fullConversation = messages;
            const structured = messages.map((m, idx) => `--- ������ #${idx + 1} ---\\n����������: ${m.sender}\\n����: ${m.subject}\\n�����������:\\n${m.body}\\n`).join("\\n");
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
        emailContext.text = "��ᢣ� ��ᚤਫ਼�.";
        finishLoading();
        return;
    }
    
    item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.trim().length > 0) {
            stopLoadingAnim();
            emailContext.text = `--- ������ ������ ---\\n����������: ${emailContext.meta.senderName}\\n����: ${emailContext.meta.subject}\\n\\n${result.value}`;
            finishLoading();
        } else {
            if (retryCount < 3) {
                setTimeout(() => fallbackCurrentMail(retryCount + 1), 500);
            } else {
                stopLoadingAnim();
                emailContext.text = `--- ������ ������ ---\\n����������: ${emailContext.meta.senderName}\\n����: ${emailContext.meta.subject}\\n\\n[��� �����婫��� ��壜�� ��� email]`;
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
// 3. EMAIL AUDIT & SUMMARY ENGINE
// -------------------------------------------------------------------------
async function generateSummaryAndAudit() {
    if (!config.apiKey) {
        document.getElementById('summaryText').innerText = '?? �������� �� API Key ���� ����婜��.';
        return;
    }

    const isPublicDomain = ['outlook.com', 'gmail.com', 'hotmail.com', 'yahoo.com', 'live.com'].includes(emailContext.myDomain);
    const currentTimeContext = new Date().toLocaleString('el-GR', { timeZone: 'Europe/Athens' });
    const currentIsoContext = new Date().toISOString();

    const prompt = `[STRICT AUTOMATION SYSTEM COMMAND]
You are a dry corporate automation API. You must parse the email thread and output EXACTLY one valid JSON object matching the JSON Schema below. No markdown, no commentary.

[CURRENT TIME CONTEXT (CRITICAL FOR CALENDAR TRACKING)]
- Today is exactly: ${currentTimeContext} (ISO: ${currentIsoContext})
Use this baseline to accurately compute relative expressions in the text like "�稠�", "��� ��棜�� ��嫞", or implicit months like "stis 20/09" -> year is 2026.

[JSON SCHEMA REQUIREMENT]
{
  "summary": "��� �������� executive �礦�� (��� 4 �������) ��� ��������, �����⤞ �� ������ 筦� ��� ��ᚤਫ਼ �� ���������. ����⨜�� ����� ⩫���� �� ��������� �㤬�� ��� �� ������� ���ᜠ.",
  "category": "High Priority",
  "urgency": "High" or "Medium" or "Low",
  "sentiment": "����������⤦�" or "���⫝̸���" or "������",
  "detected_meeting_time": "�� ������ ���������⤞ 騘/���������� ��� �������� (�.�. '��嫞 ���� 11:00') ��ᯫ�� ���, ������ ���� ''",
  "calendar_event": {
     "has_meeting": true or false,
     "label_text": "Label ��� �� UI Button (�.�. '��������ਫ਼ 20/09' � '���ᤫ��� �稠�')",
     "start_iso": "YYYY-MM-DDTHH:MM:SS format reflecting the extracted time context",
     "end_iso": "YYYY-MM-DDTHH:MM:SS format (default to 1 hour after start if unspecified)",
     "subject": "Flawless brief Greek appointment title (e.g. '���ᤫ���: [�� ���������] - [�⣘]')",
     "ai_notes": "뤘 ���⣦���, ������, �����⤦ executive �㤬�� ��� �������� ��� �����坜� ��� ������⨜��� ��� ���ᤫ���� ��� ��� ������圪. ������������ � ��㩞 raw email headers, dates, � unformatted text logs."
  },
  "smart_buttons": [
     {"label": "��������ਫ਼ 11:00", "reply_instruction": "��ᯜ ��ᤫ��� ��������ਫ਼�: ������棘��� �� �������� ��� ��� ��嫞 ���� 11:00 �������. ��� ��ᯜ�� �����梦�� ��壜��."},
     {"label": "�㫘 ������ 𨘪", "reply_instruction": "��ᯜ ��ᤫ��� 櫠 � 騘 11:00 ��� ����眠 ��� ��櫜��� �����������."}
  ]
}

[IDENTITY CONTEXT]
- Active User Name: "${emailContext.myName}"
- Active User Email: "${emailContext.myEmail}"
- Active User Corporate Domain: "@${emailContext.myDomain}"
- Public Provider Flag: ${isPublicDomain}

[CRITICAL USER MEMORY & FILTERS]
${config.agentMemory || 'None provided.'}

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
        
        document.getElementById('summaryText').innerHTML = `<p>${resObj.summary}</p>`;
        
        renderEnhancedBadges(resObj);
        renderDynamicSmartButtons(resObj.smart_buttons, resObj.calendar_event);
        updateAuditMetrics(resObj.category || 'Spam');

        if (config.autoRead) {
            setTimeout(() => {
                speakSummary();
            }, 300);
        }

        // ?? FIX: Defensive DOM Layout Measurement
        // �����⧦��� ��� ��壜�� �� ���������� ��� ����ᣜ �� ������ᜠ �� 96px (max-h-24)
        setTimeout(() => {
            const summaryContent = document.getElementById('summaryContent');
            const expandSummaryBtn = document.getElementById('expandSummaryBtn');
            const summaryFade = document.getElementById('summaryFade');
            
            if (summaryContent && summaryContent.scrollHeight > 96) {
                summaryContent.classList.add('max-h-24');
                if (expandSummaryBtn) expandSummaryBtn.classList.remove('hidden');
                if (summaryFade) summaryFade.classList.remove('hidden');
                if (expandSummaryBtn) expandSummaryBtn.innerText = '������櫜��...';
            } else {
                if (summaryContent) summaryContent.classList.remove('max-h-24');
                if (expandSummaryBtn) expandSummaryBtn.classList.add('hidden');
                if (summaryFade) summaryFade.classList.add('hidden');
            }
        }, 50);

    } catch (e) {
        console.error("? Summary Error:", e);
        document.getElementById('summaryText').innerText = '�������� ���棘��� �樫ਫ਼� �礦���.';
        document.getElementById('voiceStatus').innerText = `?? ��ᢣ� �礦���: ${e.message}`;
    }
}

function renderEnhancedBadges(resObj) {
    const badge = document.getElementById('emailCategoryBadge');
    badge.className = "text-[10px] font-bold px-2 py-0.5 rounded-full transition-all duration-300 flex items-center gap-1";
    let radarText = '';
    
    switch(resObj.category) {
        case 'High Priority':
            radarText = '?? ����� ��������櫞��';
            badge.classList.add('bg-red-500/20', 'text-red-400');
            break;
        case 'Internal':
            radarText = '?? ��૜���� / ��������';
            badge.classList.add('bg-blue-500/20', 'text-blue-400');
            break;
        case 'Newsletter':
            radarText = '?? Newsletter';
            badge.classList.add('bg-yellow-500/20', 'text-yellow-400');
            break;
        default:
            radarText = '??? ������ �������';
            badge.classList.add('bg-zinc-800', 'text-zinc-400');
    }

    if (resObj.urgency === 'High' && resObj.category !== 'Spam') radarText += ' | ?? �������';
    if (resObj.sentiment === '����������⤦�') radarText += ' | ?? �����⩡���';
    
    badge.innerText = radarText;
}

function renderDynamicSmartButtons(buttonsArray, calendarEventObj) {
    const container = document.getElementById('dynamicActionsContainer');
    if (!container) return;

    container.innerHTML = ''; 
    container.classList.remove('hidden');

    const safeButtons = (buttonsArray && buttonsArray.length > 0) ? buttonsArray : [];

    // 1. 륬��� AI �������
    safeButtons.forEach(btn => {
        const nativeBtn = document.createElement('button');
        nativeBtn.className = "flex-shrink-0 bg-secondary hover:bg-border border border-border text-xs py-1.5 px-3 rounded-full transition-colors text-primary font-medium shadow-sm flex items-center gap-1";
        nativeBtn.innerText = "? " + btn.label;
        nativeBtn.onclick = () => handleQuickAction(btn.reply_instruction);
        container.appendChild(nativeBtn);
    });

    // 2. ?? NATIVE OUTLOOK CALENDAR INTEGRATION
    if (calendarEventObj && calendarEventObj.has_meeting === true) {
        const calBtn = document.createElement('button');
        calBtn.className = "flex-shrink-0 bg-green-600/20 hover:bg-green-600/40 border border-green-500/40 text-xs py-1.5 px-3 rounded-full transition-colors text-green-400 font-bold shadow-sm flex items-center gap-1";
        calBtn.innerHTML = `?? + ������暠� (${calendarEventObj.label_text || '��������'})`;
        
        calBtn.onclick = () => {
            let startDate = calendarEventObj.start_iso ? new Date(calendarEventObj.start_iso) : new Date();
            let endDate = calendarEventObj.end_iso ? new Date(calendarEventObj.end_iso) : new Date(startDate.getTime() + 60*60*1000);

            Office.context.mailbox.displayNewAppointmentForm({
                subject: calendarEventObj.subject || `���ᤫ���: ${emailContext.meta.subject}`,
                start: startDate,
                end: endDate,
                location: calendarEventObj.location || "",
                body: calendarEventObj.ai_notes || "��������������� ���� AI Assistant.",
                requiredAttendees: [emailContext.meta.senderEmail] 
            });
        };
        container.appendChild(calBtn);
    }

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
// AUDIT DASHBOARD CONTROLS
// ----------------------
function openAuditDashboard() {
    navigate('view-audit');
    const data = JSON.parse(localStorage.getItem('emailAuditStore')) || { total: 0, high: 0, internal: 0, news: 0, spam: 0 };
    document.getElementById('auditTotal').innerText = data.total;
    document.getElementById('auditSavedTime').innerText = (data.total * 3) + '�';
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

// ?? FIX: ������, ����������� Toggle ����� syntax crashes
document.getElementById('expandSummaryBtn')?.addEventListener('click', function() {
    const summaryEl = document.getElementById('summaryContent');
    const fadeEl = document.getElementById('summaryFade');
    if (!summaryEl) return;

    if (summaryEl.classList.contains('max-h-24')) {
        summaryEl.classList.remove('max-h-24');
        if (fadeEl) fadeEl.classList.add('hidden');
        this.innerText = '���櫜��...';
    } else {
        summaryEl.classList.add('max-h-24');
        if (fadeEl) fadeEl.classList.remove('hidden');
        this.innerText = '������櫜��...';
    }
});

// -------------------------------------------------------------------------
// 4. DRAFT GENERATION ENGINE (HYPER-STRICT DIRECT FORCE)
// -------------------------------------------------------------------------
async function generateDraft(instruction, audioObj) {
    if (!config.apiKey) { navigate('view-settings'); return; }

    if (!emailContext.text || emailContext.text.trim() === "" || emailContext.text.includes("��ᢣ� ��ᚤਫ਼�.")) {
        document.getElementById('voiceStatus').innerText = "? �� email ����餜� ��棘... �����������㩫�.";
        return;
    }

    document.getElementById('voiceStatus').innerText = '?? ��⭫����...';

    const optimizedContext = emailContext.text.length > 7000
        ? emailContext.text.substring(0, 7000)
        : emailContext.text;

    const isPublicDomain = ['outlook.com', 'gmail.com', 'hotmail.com', 'yahoo.com', 'live.com'].includes(emailContext.myDomain);

    const systemPrompt = `[CRITICAL COMPILER DIRECTIVE - EXECUTE IMMEDIATELY]
You are a raw software API endpoint. You must output EXACTLY one raw valid JSON object. 
You are strictly prohibited from typing conversational preambles, general corporate holding templates, or multi-paragraph fluff.

[JSON MAPPING EXCLUSIVITY]
{"intent": "draft", "content": "YOUR_DIRECT_GREEK_RESPONSE"}

[THE ABSOLUTE COMPLIANCE LAW]
- Your single most important goal is to write a short, highly-specific email body that directly implements the [TARGET COMMAND].
- CRITICAL: If the [TARGET COMMAND] specifies a time or a clear acceptance (e.g. '��������ਫ਼ 11:00' or '������棘���'), your text MUST clearly write the confirmation of that exact fact. 
- DO NOT say "I am reviewing your data and will get back to you". The user clicked a definitive confirmation button, so you must write a final definitive response stating the confirmation!
- Keep it elegant, executive, and exactly 2-4 sentences max. Sign off using the user's name: "${emailContext.myName}".

[IDENTITY CONTEXT]
- Active User Identity Name: "${emailContext.myName}"
- Active User Email: "${emailContext.myEmail}"
- Tone Constraints: ${config.tone}
- User Strategic Memory Guidelines (Apply with priority): ${config.agentMemory || 'None.'}
- Admin Injection Filter: ${config.customPrompt || 'None.'}

[DATA LAYER - EMAIL THREAD]
${optimizedContext}

[TARGET COMMAND (YOU MUST EXECUTE THIS SPECIFIC ACTION NOW)]
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
                generationConfig: { temperature: 0.15 }
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
            parsed.content = `�����⨘ ���,\\n\\n����������� �� ���ᤫ��� ��� ��� ��� �������������⤞ 騘. �� �壘� ����⩠��� ��������.\\n\\n�� ���壞��,\\n${emailContext.myName}`;
        }

        if (parsed.intent === 'question') {
            document.getElementById('answerText').innerText = parsed.content;
            navigate('view-answer');
        } else {
            document.getElementById('draftTextarea').value = parsed.content;
            navigate('view-draft');
        }
        document.getElementById('voiceStatus').innerText = '�ᤫ� ���� ��� ������';
    } catch (e) {
        console.error("?? AI Draft Error:", e);
        document.getElementById('voiceStatus').innerText = '? ��ᢣ� AI: ' + e.message;
    }
}

// ----------------------
// ACTION EVENT BINDINGS
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
    generateDraft(`�������垩� �� ������磜�� draft email �磭घ ��: "${tweak}"\\n\\n����� EMAIL:\\n${current}`, null);
};

// ----------------------
// VOICE CAPTURE LAYER
// ----------------------
// -------------------------------------------------------------------------
// 5. VOICE CAPTURE & WAVEFORM VISUALIZATION LAYER
// -------------------------------------------------------------------------
const voiceBtn = document.getElementById('voiceBtn');
const voiceStatus = document.getElementById('voiceStatus');
const canvas = document.getElementById('waveformCanvas');

// Speech Recognition setup (Web Speech API)
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
const recognitionSupported = !!SpeechRecognition;
let recognition = null;
let recognitionText = '';

if (recognitionSupported) {
    recognition = new SpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = 'el-GR';

    recognition.onresult = (event) => {
        let interimTranscript = '';
        let finalTranscript = '';
        for (let i = event.resultIndex; i < event.results.length; ++i) {
            if (event.results[i].isFinal) {
                finalTranscript += event.results[i][0].transcript;
            } else {
                interimTranscript += event.results[i][0].transcript;
            }
        }
        
        recognitionText = finalTranscript || interimTranscript;
        
        // Dynamically insert transcribed text to active view input fields
        const draftView = document.getElementById('view-draft');
        if (draftView && draftView.classList.contains('view-active')) {
            document.getElementById('tweakPrompt').value = recognitionText;
        } else {
            document.getElementById('textPrompt').value = recognitionText;
        }
    };

    recognition.onerror = (e) => {
        console.error("Speech recognition error:", e);
    };
}

let audioContext = null;
let analyser = null;
let dataArray = null;
let bufferLength = 0;
let animationFrameId = null;

voiceBtn.onclick = () => {
    if (isRecording) { stopRecording(); return; }
    
    // Request permission from the Outlook host environment to access the microphone inside the iframe
    if (typeof Office !== 'undefined' && Office.context && Office.context.requirements.isSetSupported('DevicePermission', '1.1') && Office.devicePermission) {
        Office.devicePermission.requestPermissionsAsync([Office.DevicePermissionType.microphone], (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Microphone permission denied: " + asyncResult.error.message);
                voiceStatus.innerText = "? �������� ��橙���� ��� ������द";
            } else {
                console.log("Microphone permission verified.");
                startRecording();
            }
        });
    } else {
        // Fallback for browser testing or older Office hosts
        startRecording();
    }
};

function startRecording() {
    recognitionText = '';
    
    // Clear inputs in active panels
    const draftView = document.getElementById('view-draft');
    if (draftView && draftView.classList.contains('view-active')) {
        document.getElementById('tweakPrompt').value = '';
    } else {
        document.getElementById('textPrompt').value = '';
    }

    navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        // Start web speech recognition
        if (recognitionSupported && recognition) {
            try {
                recognition.start();
            } catch (e) {
                console.error("Failed to start speech recognition:", e);
            }
        }

        // Setup real-time waveform visualizer
        try {
            audioContext = new (window.AudioContext || window.webkitAudioContext)();
            const source = audioContext.createMediaStreamSource(stream);
            analyser = audioContext.createAnalyser();
            analyser.fftSize = 64;
            bufferLength = analyser.frequencyBinCount;
            dataArray = new Uint8Array(bufferLength);
            source.connect(analyser);

            if (canvas) {
                canvas.classList.remove('hidden');
                drawWaveform();
            }
        } catch (err) {
            console.error("Audio visualizer failed to initialize:", err);
        }

        // Fallback file recorder
        mediaRecorder = new MediaRecorder(stream, { mimeType: 'audio/webm' });
        audioChunks = [];
        mediaRecorder.ondataavailable = e => { if (e.data.size > 0) audioChunks.push(e.data); };
        
        mediaRecorder.onstop = () => {
            if (canvas) canvas.classList.add('hidden');
            if (animationFrameId) cancelAnimationFrame(animationFrameId);
            
            if (audioContext && audioContext.state !== 'closed') {
                audioContext.close();
            }

            voiceStatus.innerText = '?? �����������...';

            if (recognitionText.trim().length > 0) {
                const draftView = document.getElementById('view-draft');
                if (draftView && draftView.classList.contains('view-active')) {
                    const currentDraft = document.getElementById('draftTextarea').value;
                    generateDraft(`�������垩� �� ������磜�� draft email �磭घ ��: "${recognitionText}"\\n\\n����� EMAIL:\\n${currentDraft}`, null);
                } else {
                    generateDraft(recognitionText, null);
                }
            } else {
                const blob = new Blob(audioChunks, { type: 'audio/webm' });
                const reader = new FileReader();
                reader.readAsDataURL(blob);
                reader.onloadend = () => {
                    const base64 = reader.result.split(',')[1];
                    const draftView = document.getElementById('view-draft');
                    if (draftView && draftView.classList.contains('view-active')) {
                        const currentDraft = document.getElementById('draftTextarea').value;
                        generateDraft(`�������垩� �� ������磜�� draft email �磭घ �� ��� ���������⤞ ������. ����� EMAIL:\\n${currentDraft}`, { data: base64, mimeType: 'audio/webm' });
                    } else {
                        generateDraft('?? �ञ���� ������ ��㩫�', { data: base64, mimeType: 'audio/webm' });
                    }
                };
            }
        };

        mediaRecorder.start(1000);
        isRecording = true;
        
        voiceBtn.className = "w-24 h-24 rounded-full siri-listening flex items-center justify-center text-primary cursor-pointer border border-border";
        voiceBtn.innerHTML = `<i data-lucide="square" class="w-8 h-8 text-white"></i>`;
        lucide.createIcons();
        voiceStatus.innerText = '?? �����... ���㩫� �騘';
    }).catch(() => { 
        voiceStatus.innerText = '? ��ᢣ� ������餦�'; 
    });
}

function stopRecording() {
    if (recognitionSupported && recognition) {
        try {
            recognition.stop();
        } catch(e) {}
    }

    if (mediaRecorder && mediaRecorder.state !== 'inactive') {
        mediaRecorder.stop();
        mediaRecorder.stream.getTracks().forEach(track => track.stop());
    }
    
    isRecording = false;
    voiceBtn.className = "w-24 h-24 rounded-full siri-idle flex items-center justify-center text-primary cursor-pointer border border-border";
    voiceBtn.innerHTML = `<i data-lucide="mic" class="w-8 h-8 opacity-70"></i>`;
    lucide.createIcons();
    voiceStatus.innerText = '�ᤫ� ���� ��� ������';
}

function drawWaveform() {
    if (!isRecording) return;
    animationFrameId = requestAnimationFrame(drawWaveform);

    if (analyser && dataArray) {
        analyser.getByteFrequencyData(dataArray);
    }

    const ctx = canvas.getContext('2d');
    const width = canvas.width;
    const height = canvas.height;
    ctx.clearRect(0, 0, width, height);

    let sum = 0;
    if (dataArray) {
        for (let i = 0; i < bufferLength; i++) {
            sum += dataArray[i];
        }
    }
    const average = bufferLength > 0 ? (sum / bufferLength) : 0;
    const amplitude = Math.max(3, (average / 255) * (height * 1.1));

    const layers = [
        { color: 'rgba(59, 130, 246, 0.8)', speed: 0.02, phase: 0, width: 2.5 },
        { color: 'rgba(147, 51, 234, 0.55)', speed: 0.03, phase: 2, width: 1.5 },
        { color: 'rgba(236, 72, 153, 0.35)', speed: 0.015, phase: 4, width: 1.0 }
    ];

    layers.forEach(layer => {
        ctx.lineWidth = layer.width;
        ctx.strokeStyle = layer.color;
        ctx.beginPath();

        let x = 0;
        const sliceWidth = width / 60;

        for (let i = 0; i <= 60; i++) {
            const taper = Math.sin((i / 60) * Math.PI); // Smooth envelope
            const angle = (i * 0.15) + (Date.now() * layer.speed) + layer.phase;
            const y = (height / 2) + Math.sin(angle) * amplitude * taper;

            if (i === 0) {
                ctx.moveTo(x, y);
            } else {
                ctx.lineTo(x, y);
            }
            x += sliceWidth;
        }
        ctx.stroke();
    });
}

// ----------------------
// INSERT TO OUTLOOK
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
            voiceStatus.innerText = "?? �������� ���棘��� ����梢����.";
        }
    });
};

document.getElementById('cancelDraftBtn')?.addEventListener('click', () => {
    document.getElementById('draftTextarea').value = '';
    navigate('view-main');
});

// -------------------------------------------------------------------------
// 6. GLOBAL KEYBOARD SHORTCUTS
// -------------------------------------------------------------------------
document.addEventListener('keydown', (e) => {
    const activeEl = document.activeElement;
    const isTyping = activeEl && (
        activeEl.tagName === 'INPUT' || 
        activeEl.tagName === 'TEXTAREA' || 
        activeEl.isContentEditable
    );

    if (e.code === 'Space' && !isTyping) {
        e.preventDefault(); // Stop scrolling the view
        if (voiceBtn) voiceBtn.click();
    }

    if (e.code === 'Escape') {
        if (isRecording) {
            stopRecording();
        }
    }
});

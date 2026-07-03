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
                console.log("🔄 Item changed - reloading context");
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
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .trim();
}

// ----------------------
// GET FULL THREAD VIA REST API
// ----------------------
async function getFullConversationViaREST() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        const convId = item.conversationId;

        if (!convId) {
            console.warn("⚠️ Δεν βρέθηκε Conversation ID");
            reject(new Error("Δεν βρέθηκε Conversation ID"));
            return;
        }

        console.log(`🔍 Fetching conversation: ${convId}`);

        // Ζητάμε Callback Token για REST API (isRest: true είναι ΚΡΙΣΙΜΟ)
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (tokenResult) => {
            if (tokenResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("❌ Token error:", tokenResult.error);
                reject(new Error("Αποτυχία λήψης token. Ελέγξτε τα δικαιώματα στο Manifest."));
                return;
            }

            const token = tokenResult.value;
            const restUrl = Office.context.mailbox.restUrl;
            
            if (!restUrl) {
                reject(new Error("Δεν βρέθηκε REST URL - πιθανό πρόβλημα έκδοσης Office.js"));
                return;
            }

            // Σωστό endpoint για Outlook REST API v2.0
            const url = `${restUrl}/v2.0/me/messages?$filter=ConversationId eq '${convId}'&$select=Sender,Subject,Body,DateTimeReceived,ConversationIndex&$orderby=DateTimeReceived asc&$top=50`;

            fetch(url, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Accept': 'application/json',
                    'prefer': 'outlook.body-content-type="text"' // Ζητάμε plain text αντί για HTML
                }
            })
            .then(res => {
                if (!res.ok) {
                    return res.text().then(errText => {
                        throw new Error(`REST API Error ${res.status}: ${errText.substring(0, 200)}`);
                    });
                }
                return res.json();
            })
            .then(data => {
                if (data && data.value && data.value.length > 0) {
                    const messages = data.value
                        .filter(m => m.Body && m.Body.Content) // Φιλτράρουμε κενά
                        .map(m => {
                            let senderName = "Άγνωστος";
                            let senderEmail = "";
                            
                            if (m.Sender && m.Sender.EmailAddress) {
                                senderName = m.Sender.EmailAddress.Name || "Άγνωστος";
                                senderEmail = m.Sender.EmailAddress.Address || "";
                            }
                            
                            const bodyContent = m.Body.ContentType === 'HTML' 
                                ? cleanHtmlToText(m.Body.Content) 
                                : m.Body.Content;
                            
                            return {
                                sender: `${senderName}${senderEmail ? ` <${senderEmail}>` : ''}`,
                                subject: m.Subject || "(Χωρίς Θέμα)",
                                body: bodyContent,
                                received: m.DateTimeReceived,
                                conversationIndex: m.ConversationIndex || ""
                            };
                        });
                    
                    console.log(`✅ Loaded ${messages.length} messages from conversation`);
                    resolve(messages);
                } else {
                    console.warn("⚠️ Conversation found but empty");
                    reject(new Error("Η συνομιλία βρέθηκε κενή ή χωρίς περιεχόμενο"));
                }
            })
            .catch(err => {
                console.error("❌ REST fetch error:", err);
                reject(err);
            });
        });
    });
}

// ----------------------
// INIT OUTLOOK DATA
// ----------------------
function initOutlookData() {
    const item = Office.context.mailbox.item;
    if (!item) {
        console.warn("⚠️ No mailbox item available");
        return;
    }

    // Βασικά metadata
    emailContext.meta = {
        senderName: item.sender?.displayName || 'Άγνωστος',
        senderEmail: item.sender?.emailAddress || '',
        subject: item.subject || '(Χωρίς θέμα)',
        receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('el-GR') : '',
        conversationId: item.conversationId || null
    };

    // Ξεκινάμε loading animation
    startLoadingAnim([
        "📡 Συνδέομαι με Outlook...",
        "🔍 Αναζητώ ιστορικό συνομιλίας...",
        "🤖 Προετοιμάζω ανάλυση..."
    ]);

    // Προσπάθεια να φέρει όλο το Thread via REST API
    getFullConversationViaREST()
        .then(messages => {
            emailContext.fullConversation = messages;
            
            // Δημιουργούμε δομημένο κείμενο για το AI
            const structured = messages
                .map((m, idx) => {
                    const dateStr = m.received ? new Date(m.received).toLocaleString('el-GR') : 'Άγνωστη';
                    return `--- ΜΗΝΥΜΑ #${idx + 1} ---
ΑΠΟΣΤΟΛΕΑΣ: ${m.sender}
ΗΜΕΡΟΜΗΝΙΑ: ${dateStr}
ΘΕΜΑ: ${m.subject}
ΠΕΡΙΕΧΟΜΕΝΟ:
${m.body}
`;
                })
                .join("\n");

            emailContext.text = structured;
            console.log(`✅ Conversation loaded. Length: ${emailContext.text.length} chars`);
            finishLoading();
        })
        .catch(err => {
            console.warn("⚠️ REST API failed, using fallback:", err.message);
            // Ενημέρωση UI για fallback
            document.getElementById('loadingText').innerText = "⚠️ Χρήση τρέχοντος μηνύματος...";
            fallbackCurrentMail();
        });
}

// ----------------------
// FALLBACK: Plan B (μόνο τρέχον μήνυμα)
// ----------------------
function fallbackCurrentMail() {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) {
        stopLoadingAnim();
        emailContext.text = "Δεν ήταν δυνατή η ανάγνωση του email.";
        finishLoading();
        return;
    }

    item.body.getAsync(Office.CoercionType.Text, { asyncContext: { trim: true } }, (result) => {
        stopLoadingAnim();
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            // Προσθέτουμε metadata για context
            emailContext.text = `--- ΤΡΕΧΟΝ ΜΗΝΥΜΑ (Fallback) ---
ΑΠΟΣΤΟΛΕΑΣ: ${emailContext.meta.senderName}
ΘΕΜΑ: ${emailContext.meta.subject}
ΗΜΕΡΟΜΗΝΙΑ: ${emailContext.meta.receivedTime}

${result.value}
`;
            console.log("✅ Fallback: Loaded current message");
            finishLoading();
        } else {
            console.error("❌ Fallback failed:", result.error);
            emailContext.text = "Σφάλμα ανάγνωσης περιεχομένου.";
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
    const overlay = document.getElementById('loadingOverlay');
    
    if (overlay) {
        overlay.style.display = 'flex';
        let i = 0;
        textEl.innerText = messages[0];
        loadingInterval = setInterval(() => {
            i = (i + 1) % messages.length;
            textEl.innerText = messages[i];
        }, 1400);
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
        generateSummary();
    }
}

// ----------------------
// SUMMARY GENERATION
// ----------------------
async function generateSummary() {
    const prompt = `Κάνε σύντομη executive σύνοψη του παρακάτω email thread (max 4 γραμμές).
Περιέγραψε: Ποιος έστειλε το τελευταίο μήνυμα, τι ζητάει, και αν υπάρχουν εκκρεμότητες.

THREAD:
${emailContext.text.substring(0, 8000)}${emailContext.text.length > 8000 ? '...' : ''}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        const res = await fetch(url, {
            method: "POST", 
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ 
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.3, maxOutputTokens: 256 }
            })
        });

        const data = await res.json();
        if (data.error) throw new Error(data.error.message);

        const summary = data.candidates?.[0]?.content?.parts?.[0]?.text || 'Δεν παρήχθη σύνοψη';
        document.getElementById('summaryText').innerText = summary;
        
        // Show expand button if text is long
        const summaryEl = document.getElementById('summaryText');
        const fadeEl = document.getElementById('summaryFade');
        const expandBtn = document.getElementById('expandSummaryBtn');
        
        if (summary.length > 150) {
            summaryEl.classList.add('max-h-24');
            fadeEl?.classList.remove('hidden');
            expandBtn?.classList.remove('hidden');
        }
    } catch (e) {
        console.error("Summary error:", e);
        document.getElementById('summaryText').innerText = '⚠️ Σφάλμα σύνοψης: ' + e.message;
    }
}

// Expand summary toggle
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

// Manual summary button
document.getElementById('manualSummaryBtn')?.addEventListener('click', () => {
    if (config.apiKey && emailContext.text) {
        generateSummary();
    } else if (!config.apiKey) {
        alert("Προσθέστε API Key στις ρυθμίσεις");
        navigate('view-settings');
    }
});

// ----------------------
// QUICK ACTIONS & TWEAKS
// ----------------------
function handleQuickAction(actionType) {
    if (!config.apiKey) { 
        alert("Προσθέστε Gemini API Key στις ρυθμίσεις πρώτα"); 
        navigate('view-settings');
        return; 
    }
    generateDraft(actionType, null);
}

document.getElementById('sendTextBtn').onclick = () => {
    const txt = document.getElementById('textPrompt').value.trim();
    if (!txt) return;
    if (!config.apiKey) {
        alert("Προσθέστε API Key πρώτα");
        navigate('view-settings');
        return;
    }
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
// VOICE RECORDING
// ----------------------
const voiceBtn = document.getElementById('voiceBtn');
const voiceStatus = document.getElementById('voiceStatus');

voiceBtn.onclick = () => {
    if (!config.apiKey) { 
        alert("Προσθέστε Gemini API Key στις ρυθμίσεις"); 
        navigate('view-settings');
        return; 
    }
    if (isRecording) { 
        stopRecording(); 
        return; 
    }

    // Request microphone permission via Office API if available
    if (Office.context.mailbox && Office.devicePermission) {
        Office.devicePermission.requestPermissionsAsync([Office.DevicePermissionType.microphone], (res) => {
            if (res.status === Office.AsyncResultStatus.Failed) {
                alert("❌ Η πρόσβαση στο μικρόφωνο απορρίφθηκε. Ελέγξτε τις ρυθμίσεις απορρήτου.");
                return;
            }
            // If permission already granted (res.value === true), reload to apply
            if (res.value) {
                window.location.reload(); 
            } else {
                startRecording();
            }
        });
    } else {
        // Fallback to browser API
        startRecording();
    }
};

function startRecording() {
    navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        mediaRecorder = new MediaRecorder(stream, { mimeType: 'audio/webm' });
        audioChunks = [];
        
        mediaRecorder.ondataavailable = e => { 
            if (e.data.size > 0) audioChunks.push(e.data); 
        };
        
        mediaRecorder.onstop = () => {
            voiceStatus.innerText = '🔄 Επεξεργασία ήχου...';
            const blob = new Blob(audioChunks, { type: 'audio/webm' });
            const reader = new FileReader();
            
            reader.readAsDataURL(blob);
            reader.onloadend = () => {
                const base64 = reader.result.split(',')[1];
                generateDraft('🎤 Φωνητική εντολή χρήστη (δείτε το audio attachment)', { 
                    data: base64, 
                    mimeType: 'audio/webm' 
                });
            };
        };

        mediaRecorder.start(1000); // chunk every second for reliability
        isRecording = true;
        
        // UI updates
        voiceBtn.classList.remove('siri-idle');
        voiceBtn.classList.add('siri-listening');
        voiceBtn.innerHTML = `<i data-lucide="square" class="w-8 h-8 text-white"></i>`;
        lucide.createIcons();
        voiceStatus.innerText = '🔴 Καταγραφή... Μιλήστε τώρα';
        
    }).catch(err => {
        console.error("🎤 Mic error:", err);
        voiceStatus.innerText = '❌ Σφάλμα μικροφώνου: ' + (err.message || 'Άγνωστο σφάλμα');
    });
}

function stopRecording() {
    if (mediaRecorder && mediaRecorder.state !== 'inactive') {
        mediaRecorder.stop();
        // Stop all tracks to release mic
        mediaRecorder.stream.getTracks().forEach(track => track.stop());
    }
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
    if (!config.apiKey) {
        alert("Προσθέστε API Key πρώτα");
        navigate('view-settings');
        return;
    }

    voiceStatus.innerText = '🤖 Σκέφτομαι...';

    // Truncate context if too long (Gemini limit ~32k tokens, but be safe)
    const contextText = emailContext.text.length > 25000 
        ? emailContext.text.substring(0, 25000) + '\n\n[...περιορισμένο λόγω μήκους...]' 
        : emailContext.text;

    const systemPrompt = `Είσαι Executive AI Assistant για επαγγελματική αλληλογραφία.
Έχεις πρόσβαση στο ΠΛΗΡΕΣ ιστορικό μιας συνομιλίας (Email Thread).
Τα μηνύματα είναι χωρισμένα με "--- ΜΗΝΥΜΑ #Χ ---". Το τελευταίο χρονολογικά βρίσκεται στο τέλος.

📧 EMAIL THREAD:
${contextText}

👤 USER INSTRUCTION:
${instruction}

🎯 ΚΑΘΗΚΟΝ:
1. Κατάλαβε αν ο χρήστης ζητάει να ΓΡΑΨΕΙΣ EMAIL απάντησης (π.χ. "απάντα", "ευχαρίστησέ τον") 
   ή αν κάνει ΕΡΩΤΗΣΗ/ΣΥΖΗΤΗΣΗ για το thread (π.χ. "τι λέει εδώ;", "ποιος είναι;").
2. Απάντησε ΑΥΣΤΗΡΑ με JSON (χωρίς markdown, χωρίς εξηγήσεις):
{
 "intent": "draft" ή "question",
 "content": "Το κείμενό σου"
}

📝 ΚΑΝΟΝΕΣ:
• Αν intent="draft": Γράψε ΕΤΟΙΜΟ επαγγελματικό email ΑΠΑΝΤΩΝΤΑΣ στο ΤΕΛΕΥΤΑΙΟ μήνυμα του thread. 
  Χρησιμοποίησε τόνο: ${config.tone}. Μην συμπεριλάβεις υπογραφή εκτός αν ζητηθεί.
• Αν intent="question": Απάντα ξεκάθαρα στην ερώτηση σε μορφή κειμένου (όχι email).
• Χρησιμοποίησε Ελληνικά εκτός αν το thread είναι σε άλλη γλώσσα.
${config.customPrompt ? `\n• ΠΡΟΣΘΕΤΗ ΟΔΗΓΙΑ: ${config.customPrompt}` : ''}`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${config.apiKey}`;
        
        const parts = [{ text: systemPrompt }];
        if (audioObj?.data) {
            parts.push({ 
                inlineData: { 
                    mimeType: audioObj.mimeType || 'audio/webm', 
                    data: audioObj.data 
                } 
            });
        }

        const res = await fetch(url, {
            method: "POST", 
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts }],
                generationConfig: { 
                    responseMimeType: "application/json",
                    temperature: 0.2,
                    maxOutputTokens: 2048
                }
            })
        });

        const data = await res.json();
        
        if (data.error) {
            throw new Error(data.error.message || 'Unknown API error');
        }

        let raw = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        if (!raw) throw new Error("Κενή απάντηση από το AI");
        
        // Clean markdown if present
        raw = raw.replace(/```json/g, '').replace(/```/g, '').trim();
        
        const parsed = JSON.parse(raw);

        if (parsed.intent === 'question') {
            document.getElementById('answerText').innerText = parsed.content;
            navigate('view-answer');
        } else {
            document.getElementById('draftTextarea').value = parsed.content;
            navigate('view-draft');
        }

        voiceStatus.innerText = '✅ Ολοκληρώθηκε!';
        
    } catch (e) {
        console.error("🤖 AI Error:", e);
        alert(`⚠️ Σφάλμα AI: ${e.message}\n\nΕλέγξτε:\n• Έγκυρο API Key\n• Σταθερή σύνδεση\n• Μήκος συνομιλίας`);
        voiceStatus.innerText = '❌ Σφάλμα - Δοκιμάστε ξανά';
    }
}

// ----------------------
// INSERT TO OUTLOOK
// ----------------------
document.getElementById('insertOutlookBtn').onclick = () => {
    const finalTxt = document.getElementById('draftTextarea').value;
    if (!finalTxt.trim()) {
        alert("Το draft είναι κενό");
        return;
    }
    
    Office.context.mailbox.item.displayReplyForm(finalTxt, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("✅ Reply form opened");
            document.getElementById('draftTextarea').value = '';
            navigate('view-main');
        } else {
            console.error("❌ displayReplyForm error:", asyncResult.error);
            alert("⚠️ Δεν ήταν δυνατό το άνοιγμα της φόρμας απάντησης. Δοκιμάστε χειροκίνητα αντιγραφή-επικόλληση.");
        }
    });
};

// Cancel draft button
document.getElementById('cancelDraftBtn')?.addEventListener('click', () => {
    document.getElementById('draftTextarea').value = '';
    navigate('view-main');
});

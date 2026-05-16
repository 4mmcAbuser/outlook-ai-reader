// Περιμένουμε να φορτώσει το Outlook
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initApp();
    }
});

let apiKey = "";
let mediaStream = null;         // για να κλείνουμε το stream μετά
let recognition = null;
let isListening = false;
let pressTimer = null;

function initApp() {
    // 1. Φόρτωση αποθηκευμένου API Key
    apiKey = localStorage.getItem("geminiApiKey");
    if (apiKey) {
        document.getElementById("apiKeyInput").value = apiKey;
        enableMicrophoneButton(true);
        document.getElementById("status").innerText = "Έτοιμο για χρήση!";
    }

    // 2. Αποθήκευση κλειδιού
    document.getElementById("saveKeyBtn").onclick = () => {
        let key = document.getElementById("apiKeyInput").value.trim();
        if (key) {
            localStorage.setItem("geminiApiKey", key);
            apiKey = key;
            enableMicrophoneButton(true);
            document.getElementById("status").innerHTML = "✅ Το κλειδί αποθηκεύτηκε!";
        }
    };

    // 3. Έλεγχος υποστήριξης SpeechRecognition
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
        document.getElementById("status").innerHTML = "❌ Το πρόγραμμα περιήγησης του Outlook δεν υποστηρίζει υπαγόρευση. Χρησιμοποιήστε πληκτρολόγηση.";
        enableMicrophoneButton(false);
        return;
    }

    recognition = new SpeechRecognition();
    recognition.lang = 'el-GR';
    recognition.interimResults = false;
    recognition.continuous = false;

    // 4. Λειτουργία "Κράτα πατημένο"
    const actionBtn = document.getElementById("actionBtn");
    
    actionBtn.addEventListener("mousedown", () => {
        if (actionBtn.disabled) return;
        // Ξεκινάμε υπαγόρευση μετά από 100ms (αποφυγή false trigger)
        pressTimer = setTimeout(() => {
            startVoiceRecording();
        }, 100);
    });
    
    actionBtn.addEventListener("mouseup", () => {
        if (pressTimer) {
            clearTimeout(pressTimer);
            pressTimer = null;
        }
        // Δεν κάνουμε stop – θα σταματήσει μόνο του όταν τελειώσει η ομιλία
        // ή αν θέλετε να διακόπτετε απότομα, μπορείτε να καλέσετε recognition.stop()
    });
    
    actionBtn.addEventListener("mouseleave", () => {
        if (pressTimer) {
            clearTimeout(pressTimer);
            pressTimer = null;
        }
        if (isListening) {
            recognition.stop();
        }
    });

    recognition.onstart = () => {
        isListening = true;
        document.getElementById("status").innerHTML = "🎤 Ακούω... Μίλα τώρα!";
        document.getElementById("status").style.color = "#ff453a";
        document.getElementById("result").style.display = "none";
    };
    
    recognition.onend = () => {
        isListening = false;
        // Κλείνουμε το stream μικροφώνου για να ελευθερωθεί η συσκευή
        if (mediaStream) {
            mediaStream.getTracks().forEach(track => track.stop());
            mediaStream = null;
        }
        document.getElementById("status").innerHTML = "✅ Έτοιμο (πατήστε παρατεταμένα)";
        document.getElementById("status").style.color = "#32d74b";
    };
    
    recognition.onresult = (event) => {
        const voiceCommand = event.results[0][0].transcript;
        document.getElementById("status").innerHTML = `📝 Είπες: "${voiceCommand}"<br>🔄 Διαβάζω το email...`;
        document.getElementById("status").style.color = "#0a84ff";
        extractEmailAndProcess(voiceCommand);
    };
    
    recognition.onerror = (event) => {
        console.error("SpeechRecognition error:", event.error);
        let errorMsg = "";
        switch(event.error) {
            case 'not-allowed':
                errorMsg = "Δεν δώσατε άδεια μικροφώνου. Παρακαλώ επιτρέψτε την πρόσβαση και δοκιμάστε ξανά.";
                break;
            case 'no-speech':
                errorMsg = "Δεν ανιχνεύθηκε ομιλία. Προσπαθήστε ξανά.";
                break;
            default:
                errorMsg = "Σφάλμα μικροφώνου: " + event.error;
        }
        document.getElementById("status").innerHTML = "❌ " + errorMsg;
        document.getElementById("status").style.color = "#ff453a";
        if (mediaStream) {
            mediaStream.getTracks().forEach(track => track.stop());
            mediaStream = null;
        }
        isListening = false;
    };
}

// Συνάρτηση που ζητάει άδεια μικροφώνου και μετά ξεκινά το recognition
async function startVoiceRecording() {
    if (!recognition) {
        document.getElementById("status").innerHTML = "❌ Η υπαγόρευση δεν υποστηρίζεται.";
        return;
    }
    if (isListening) return;
    
    // Ζητάμε δικαίωμα μικροφώνου (απαραίτητο στο add‑in WebView)
    try {
        const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
        mediaStream = stream; // κρατάμε για να κλείσουμε μετά
        recognition.start();
    } catch (err) {
        console.error("getUserMedia error:", err);
        let msg = "Δεν μπόρεσα να αποκτήσω πρόσβαση στο μικρόφωνο.";
        if (err.name === 'NotAllowedError' || err.name === 'PermissionDeniedError') {
            msg = "Η πρόσβαση στο μικρόφωνο απορρίφθηκε. Ελέγξτε τις ρυθμίσεις ασφαλείας του Outlook.";
        }
        document.getElementById("status").innerHTML = "❌ " + msg;
        document.getElementById("status").style.color = "#ff453a";
        if (mediaStream) {
            mediaStream.getTracks().forEach(track => track.stop());
            mediaStream = null;
        }
    }
}

function enableMicrophoneButton(enabled) {
    const btn = document.getElementById("actionBtn");
    btn.disabled = !enabled;
    if (!enabled) {
        btn.style.opacity = "0.5";
    } else {
        btn.style.opacity = "1";
    }
}

// Οι υπόλοιπες συναρτήσεις extractEmailAndProcess και callGeminiAPI παραμένουν IDENTICAL με πριν
// (δεν τις αλλάζουμε, αλλά επισυνάπτονται για πληρότητα)

function extractEmailAndProcess(voiceCommand) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            document.getElementById("status").innerHTML = "🤖 Στέλνω δεδομένα στο Gemini AI...";
            callGeminiAPI(emailBody, voiceCommand);
        } else {
            document.getElementById("status").innerHTML = "❌ Σφάλμα ανάγνωσης email.";
        }
    });
}

async function callGeminiAPI(emailText, voiceCommand) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    
    const prompt = `
Είσαι Executive Assistant. 
Email Πελάτη: "${emailText}"
Φωνητική Εντολή Αφεντικού: "${voiceCommand}"

Εξήγαγε JSON με: "summary" (σύνοψη), "email_reply" (επίσημη απάντηση), και "order_data" (δεδομένα φόρμας πχ όνομα, ποσό).
Απάντησε ΑΥΣΤΗΡΑ ΚΑΙ ΜΟΝΟ σε μορφή JSON.
`;

    const requestBody = {
        "contents": [{ "parts": [{"text": prompt}] }],
        "generationConfig": { "responseMimeType": "application/json" }
    };

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(requestBody)
        });

        const data = await response.json();
        
        if (data.error) throw new Error(data.error.message);

        const aiResultText = data.candidates[0].content.parts[0].text;
        
        document.getElementById("status").innerHTML = "✅ Επιτυχία!";
        document.getElementById("status").style.color = "#32d74b";
        document.getElementById("result").style.display = "block";
        document.getElementById("result").innerText = aiResultText;

    } catch (error) {
        document.getElementById("status").innerHTML = "❌ Σφάλμα AI: " + error.message;
        document.getElementById("status").style.color = "#ff453a";
    }
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initApp();
    }
});

let apiKey = "";

function initApp() {
    // Φόρτωση API Key
    apiKey = localStorage.getItem("geminiApiKey");
    if (apiKey) {
        document.getElementById("apiKeyInput").value = apiKey;
        document.getElementById("actionBtn").disabled = false;
        document.getElementById("status").innerText = "Έτοιμο για χρήση!";
    }

    // Αποθήκευση API Key
    document.getElementById("saveKeyBtn").onclick = () => {
        let key = document.getElementById("apiKeyInput").value.trim();
        if (key) {
            localStorage.setItem("geminiApiKey", key);
            apiKey = key;
            document.getElementById("actionBtn").disabled = false;
            document.getElementById("status").innerText = "Το κλειδί αποθηκεύτηκε!";
            document.getElementById("status").style.color = "#32d74b";
        }
    };

    // Το κουμπί Υπαγόρευσης
    const actionBtn = document.getElementById("actionBtn");

    actionBtn.onclick = () => {
        // 1. Ζητάμε άδεια μέσω της Microsoft (Το API που βρήκες!)
        if (Office.context.mailbox && Office.devicePermission) {
            const deviceCapabilities = [Office.DevicePermissionType.microphone];

            Office.devicePermission.requestPermissionsAsync(deviceCapabilities, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    document.getElementById("status").innerText = "Αρνηθήκατε την πρόσβαση στο μικρόφωνο.";
                    document.getElementById("status").style.color = "#ff453a";
                } else {
                    if (asyncResult.value) {
                        // Η άδεια δόθηκε για πρώτη φορά. Η Microsoft ΑΠΑΙΤΕΙ reload!
                        document.getElementById("status").innerText = "Η άδεια δόθηκε! Γίνεται ανανέωση...";
                        location.reload();
                    } else {
                        // Η άδεια υπάρχει ήδη (asyncResult.value === false). Ξεκινάμε!
                        startListening();
                    }
                }
            });
        } else {
            // Αν για κάποιο λόγο δεν υποστηρίζεται το API, δοκιμάζουμε απευθείας
            startListening();
        }
    };
}

// 2. Η συνάρτηση που ανοίγει το μικρόφωνο
function startListening() {
    let recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
    recognition.lang = 'el-GR';
    recognition.interimResults = false;

    document.getElementById("status").innerText = "Ακούω... Μίλα τώρα!";
    document.getElementById("status").style.color = "#ff453a";
    document.getElementById("result").style.display = "none";
    
    recognition.start();

    recognition.onresult = (event) => {
        const voiceCommand = event.results[0][0].transcript;
        document.getElementById("status").innerText = `Είπες: "${voiceCommand}"\nΔιαβάζω το email...`;
        document.getElementById("status").style.color = "#0a84ff";
        
        extractEmailAndProcess(voiceCommand);
    };

    recognition.onerror = (event) => {
        document.getElementById("status").innerText = "Σφάλμα μικροφώνου: " + event.error;
        document.getElementById("status").style.color = "#ff453a";
    };
}

// 3. Διαβάζει το email και καλεί το Gemini
function extractEmailAndProcess(voiceCommand) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            document.getElementById("status").innerText = "Στέλνω δεδομένα στο Gemini AI...";
            callGeminiAPI(emailBody, voiceCommand);
        } else {
            document.getElementById("status").innerText = "Σφάλμα ανάγνωσης email.";
        }
    });
}

// 4. Η κλήση στο Gemini
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
        
        document.getElementById("status").innerText = "Επιτυχία!";
        document.getElementById("status").style.color = "#32d74b";
        document.getElementById("result").style.display = "block";
        document.getElementById("result").innerText = aiResultText;

    } catch (error) {
        document.getElementById("status").innerText = "Σφάλμα AI: " + error.message;
        document.getElementById("status").style.color = "#ff453a";
    }
}

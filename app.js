// Περιμένουμε να φορτώσει το Outlook
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initApp();
    }
});

let apiKey = "";

function initApp() {
    // 1. Έλεγχος αν υπάρχει ήδη αποθηκευμένο API Key
    apiKey = localStorage.getItem("geminiApiKey");
    if (apiKey) {
        document.getElementById("apiKeyInput").value = apiKey;
        document.getElementById("actionBtn").disabled = false;
        document.getElementById("status").innerText = "Έτοιμο για χρήση!";
    }

    // 2. Αποθήκευση του API Key
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

    // 3. Λειτουργία Υπαγόρευσης (STT) + Επεξεργασία
    const actionBtn = document.getElementById("actionBtn");
    let recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
    recognition.lang = 'el-GR'; // Ελληνική Αναγνώριση
    recognition.interimResults = false;

    // Όταν ο χρήστης πατήσει το κουμπί
    actionBtn.onclick = () => {
        document.getElementById("status").innerText = "Ακούω... Μίλα τώρα!";
        document.getElementById("status").style.color = "#ff453a";
        document.getElementById("result").style.display = "none";
        recognition.start();
    };

    // Όταν τελειώσει η ομιλία
    recognition.onresult = (event) => {
        const voiceCommand = event.results[0][0].transcript;
        document.getElementById("status").innerText = `Είπες: "${voiceCommand}"\nΔιαβάζω το email...`;
        document.getElementById("status").style.color = "#0a84ff";
        
        // Διαβάζουμε το Email
        extractEmailAndProcess(voiceCommand);
    };

    recognition.onerror = (event) => {
        document.getElementById("status").innerText = "Σφάλμα μικροφώνου: " + event.error;
        document.getElementById("status").style.color = "#ff453a";
    };
}

function extractEmailAndProcess(voiceCommand) {
    // Παίρνουμε το κείμενο του Email από το Outlook
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            document.getElementById("status").innerText = "Στέλνω δεδομένα στο Gemini AI...";
            
            // Κλήση στο Gemini AI
            callGeminiAPI(emailBody, voiceCommand);
        } else {
            document.getElementById("status").innerText = "Σφάλμα ανάγνωσης email.";
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
        
        // Εμφάνιση του αποτελέσματος στην οθόνη
        document.getElementById("status").innerText = "Επιτυχία!";
        document.getElementById("status").style.color = "#32d74b";
        document.getElementById("result").style.display = "block";
        document.getElementById("result").innerText = aiResultText;

    } catch (error) {
        document.getElementById("status").innerText = "Σφάλμα AI: " + error.message;
        document.getElementById("status").style.color = "#ff453a";
    }
}
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initApp();
    }
});

let apiKey = "";
let isRecording = false;
let mediaRecorder;
let audioChunks = [];

function initApp() {
    apiKey = localStorage.getItem("geminiApiKey");
    if (apiKey) {
        document.getElementById("apiKeyInput").value = apiKey;
        document.getElementById("actionBtn").disabled = false;
        document.getElementById("status").innerText = "Έτοιμο για χρήση!";
    }

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

    const actionBtn = document.getElementById("actionBtn");

    actionBtn.onclick = () => {
        if (isRecording) {
            // Αν γράφει ήδη, το σταματάμε
            stopRecording();
            return;
        }

        // Ζητάμε άδεια μικροφώνου (η γνωστή γραφειοκρατία του Outlook)
        if (Office.context.mailbox && Office.devicePermission) {
            Office.devicePermission.requestPermissionsAsync([Office.DevicePermissionType.microphone], (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    document.getElementById("status").innerText = "Αρνηθήκατε την πρόσβαση.";
                } else {
                    if (asyncResult.value) {
                        location.reload(); // Πρώτη φορά -> Reload
                    } else {
                        startRecording(); // Έχει ήδη άδεια -> Πάμε!
                    }
                }
            });
        } else {
            startRecording();
        }
    };
}

// 1. ΗΧΟΓΡΑΦΗΣΗ ΦΩΝΗΣ (Αντί για STT)
function startRecording() {
    navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        mediaRecorder = new MediaRecorder(stream);
        audioChunks = [];
        
        mediaRecorder.ondataavailable = e => {
            if (e.data.size > 0) audioChunks.push(e.data);
        };

        mediaRecorder.onstop = () => {
            document.getElementById("status").innerText = "Ηχογράφηση ολοκληρώθηκε. Διαβάζω Email...";
            
            // Πακετάρουμε τον ήχο
            const audioBlob = new Blob(audioChunks, { type: mediaRecorder.mimeType || 'audio/webm' });
            
            // Μετατροπή σε Base64 για να πάει στο API
            const reader = new FileReader();
            reader.readAsDataURL(audioBlob);
            reader.onloadend = () => {
                const base64Audio = reader.result.split(',')[1];
                const cleanMimeType = (mediaRecorder.mimeType || 'audio/webm').split(';')[0];
                extractEmailAndProcess(base64Audio, cleanMimeType);
            };
        };

        mediaRecorder.start();
        isRecording = true;
        
        // Αλλάζουμε το UI
        const actionBtn = document.getElementById("actionBtn");
        actionBtn.innerText = "🛑 Πάτα για Τερματισμό";
        actionBtn.style.background = "#ff453a";
        document.getElementById("status").innerText = "🔴 Ηχογράφηση... Μίλα τώρα!";
        document.getElementById("status").style.color = "#ff453a";
        document.getElementById("result").style.display = "none";

    }).catch(err => {
        document.getElementById("status").innerText = "Σφάλμα μικροφώνου: Δεν βρέθηκε συσκευή.";
    });
}

function stopRecording() {
    if (mediaRecorder && mediaRecorder.state !== "inactive") {
        mediaRecorder.stop();
    }
    isRecording = false;
    const actionBtn = document.getElementById("actionBtn");
    actionBtn.innerText = "🎤 Πάτα για Ηχογράφηση";
    actionBtn.style.background = "#32d74b";
}

// 2. ΔΙΑΒΑΣΜΑ EMAIL
function extractEmailAndProcess(base64Audio, mimeType) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            document.getElementById("status").innerText = "Στέλνω Ήχο και Email στο AI...";
            document.getElementById("status").style.color = "#0a84ff";
            callGeminiAudioAPI(emailBody, base64Audio, mimeType);
        } else {
            document.getElementById("status").innerText = "Σφάλμα ανάγνωσης email.";
        }
    });
}

// 3. Η ΜΑΓΕΙΑ: Στέλνουμε Ήχο και Κείμενο μαζί στο Gemini 1.5 Flash
async function callGeminiAudioAPI(emailText, base64Audio, mimeType) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;
    
    const prompt = `Είσαι Executive Assistant. 
Email Πελάτη: "${emailText}"

Άκουσε την ηχητική εντολή του αφεντικού από το επισυναπτόμενο αρχείο ήχου. Μπορεί να είναι στα Ελληνικά, Αγγλικά ή Greeklish.
Εξήγαγε JSON με: "summary" (σύνοψη), "email_reply" (επίσημη απάντηση), και "order_data" (δεδομένα φόρμας πχ όνομα, ποσό).
Απάντησε ΑΥΣΤΗΡΑ ΚΑΙ ΜΟΝΟ σε μορφή JSON.`;

    const requestBody = {
        "contents": [
            { 
                "parts": [
                    { "text": prompt },
                    {
                        "inlineData": {
                            "mimeType": mimeType,
                            "data": base64Audio
                        }
                    }
                ] 
            }
        ],
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

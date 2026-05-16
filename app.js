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

    // ΝΕΟ: Τι γίνεται όταν πατάς "Επικόλληση στο Outlook"
    document.getElementById("insertReplyBtn").onclick = () => {
        const finalEmailText = document.getElementById("draftText").value;
        
        // Η ΕΝΤΟΛΗ ΤΗΣ MICROSOFT ΠΟΥ ΑΝΟΙΓΕΙ ΤΗΝ ΑΠΑΝΤΗΣΗ!
        Office.context.mailbox.item.displayReplyForm(finalEmailText);
        
        document.getElementById("status").innerText = "✅ Η απάντηση άνοιξε στο Outlook!";
        document.getElementById("status").style.color = "#32d74b";
    };

    const actionBtn = document.getElementById("actionBtn");

    actionBtn.onclick = () => {
        if (isRecording) {
            stopRecording();
            return;
        }

        if (Office.context.mailbox && Office.devicePermission) {
            Office.devicePermission.requestPermissionsAsync([Office.DevicePermissionType.microphone], (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    document.getElementById("status").innerText = "Αρνηθήκατε την πρόσβαση.";
                } else {
                    if (asyncResult.value) {
                        location.reload(); 
                    } else {
                        startRecording(); 
                    }
                }
            });
        } else {
            startRecording();
        }
    };
}

function startRecording() {
    navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        mediaRecorder = new MediaRecorder(stream);
        audioChunks = [];
        
        mediaRecorder.ondataavailable = e => {
            if (e.data.size > 0) audioChunks.push(e.data);
        };

        mediaRecorder.onstop = () => {
            document.getElementById("status").innerText = "Ηχογράφηση ολοκληρώθηκε. Διαβάζω Ιστορικό...";
            const audioBlob = new Blob(audioChunks, { type: mediaRecorder.mimeType || 'audio/webm' });
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
        
        const actionBtn = document.getElementById("actionBtn");
        actionBtn.innerText = "🛑 Πάτα για Τερματισμό";
        actionBtn.style.background = "#ff453a";
        document.getElementById("status").innerText = "🔴 Ηχογράφηση... Μίλα τώρα!";
        document.getElementById("status").style.color = "#ff453a";
        document.getElementById("result").style.display = "none";
        document.getElementById("draftContainer").style.display = "none";

    }).catch(err => {
        document.getElementById("status").innerText = "Σφάλμα μικροφώνου.";
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

function extractEmailAndProcess(base64Audio, mimeType) {
    // Το getAsync τραβάει ΟΛΟ το ιστορικό που φαίνεται στο email
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            document.getElementById("status").innerText = "Στέλνω δεδομένα στο AI...";
            document.getElementById("status").style.color = "#0a84ff";
            callGeminiAudioAPI(emailBody, base64Audio, mimeType);
        } else {
            document.getElementById("status").innerText = "Σφάλμα ανάγνωσης email.";
        }
    });
}

async function callGeminiAudioAPI(emailText, base64Audio, mimeType) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
    
    const prompt = `Είσαι Executive Assistant. 
Ιστορικό Συνομιλίας Πελάτη: "${emailText}"

Άκουσε την ηχητική εντολή του αφεντικού.
Εξήγαγε JSON με: "summary" (σύνοψη), "email_reply" (επίσημη απάντηση στο email του πελάτη), και "order_data" (αν ζητήθηκε φόρμα/εντολή).
Απάντησε ΑΥΣΤΗΡΑ σε μορφή JSON, χωρίς μορφοποίηση markdown στην αρχή (π.χ. χωρίς \`\`\`json).`;

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
        
        // Μετατρέπουμε το Κείμενο σε πραγματικό JSON (Dictionary)
        const parsedData = JSON.parse(aiResultText);
        
        // Βάζουμε τη Σύνοψη και τη Φόρμα στο μαύρο κουτί (Για το δικό σου σύστημα)
        document.getElementById("result").innerText = "Σύνοψη: " + parsedData.summary + "\n\nΔεδομένα Φόρμας: \n" + JSON.stringify(parsedData.order_data, null, 2);
        document.getElementById("result").style.display = "block";

        // Βάζουμε την Απάντηση στο νέο Textarea για να τη δει/διορθώσει ο χρήστης
        document.getElementById("draftText").value = parsedData.email_reply;
        document.getElementById("draftContainer").style.display = "block";

        document.getElementById("status").innerText = "✨ Το AI ολοκλήρωσε!";
        document.getElementById("status").style.color = "#32d74b";

    } catch (error) {
        document.getElementById("status").innerText = "Σφάλμα AI: " + error.message;
        document.getElementById("status").style.color = "#ff453a";
    }
}

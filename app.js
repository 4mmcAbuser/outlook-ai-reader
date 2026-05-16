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

    document.getElementById("insertReplyBtn").onclick = () => {
        const finalEmailText = document.getElementById("draftText").value;
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
            document.getElementById("status").innerText = "Ηχογράφηση ολοκληρώθηκε. Συλλέγω δεδομένα συνομιλίας...";
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
        
        document.getElementById("actionBtn").innerText = "🛑 Πάτα για Τερματισμό";
        document.getElementById("actionBtn").style.background = "#ff453a";
        document.getElementById("status").innerText = "🔴 Ηχογράφηση... Μίλα τώρα!";
        document.getElementById("status").style.color = "#ff453a";
        
        // Κρύβουμε τα παλιά αποτελέσματα
        document.getElementById("voiceInputContainer").style.display = "none";
        document.getElementById("formContainer").style.display = "none";
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
    document.getElementById("actionBtn").innerText = "🎤 Πάτα για Ηχογράφηση";
    document.getElementById("actionBtn").style.background = "#32d74b";
}

function extractEmailAndProcess(base64Audio, mimeType) {
    const item = Office.context.mailbox.item;
    
    // ΝΕΟ: Μαζεύουμε τα στοιχεία "Ποιος μιλάει και σε ποιον"
    const emailMetadata = {
        senderName: item.sender ? item.sender.displayName : "Άγνωστος",
        senderEmail: item.sender ? item.sender.emailAddress : "unknown@mail.com",
        subject: item.subject,
        toRecipients: item.to ? item.to.map(r => `${r.displayName} (${r.emailAddress})`).join(", ") : "Μόνο εγώ"
    };

    // Διαβάζουμε όλο το ιστορικό της συνομιλίας (body)
    item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            document.getElementById("status").innerText = "Αναλύω τη συνομιλία με το Gemini 2.5...";
            document.getElementById("status").style.color = "#0a84ff";
            
            callGeminiAudioAPI(emailBody, emailMetadata, base64Audio, mimeType);
        } else {
            document.getElementById("status").innerText = "Σφάλμα ανάγνωσης συνομιλίας.";
        }
    });
}

async function callGeminiAudioAPI(emailText, meta, base64Audio, mimeType) {
    // Χρήση του gemini-2.5-flash-lite που πρότεινες!
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;
    
    const prompt = `Είσαι ο κορυφαίος Executive Assistant της εταιρείας.
Ανάλυσε την παρακάτω συνομιλία και τα στοιχεία της για να καταλάβεις ΠΟΙΟΣ μιλάει, σε ΠΟΙΟΝ και ΓΙΑΤΙ (ποιο είναι το θέμα/πρόθεση).

--- ΣΤΟΙΧΕΙΑ ΜΗΝΥΜΑΤΟΣ ---
Αποστολέας (Αυτός που μας έστειλε το mail): ${meta.senderName} (${meta.senderEmail})
Θέμα Συζήτησης: ${meta.subject}
Παραλήπτες: ${meta.toRecipients}

--- ΟΛΟΚΛΗΡΟ ΤΟ ΙΣΤΟΡΙΚΟ ΤΗΣ ΣΥΝΟΜΙΛΙΑΣ (THREAD) ---
"${emailText}"

--- ΟΔΗΓΙΕΣ ΦΩΝΗΣ ---
Άκουσε το αρχείο ήχου. Περιέχει τη φωνητική εντολή του αφεντικού για το τι πρέπει να απαντήσουμε ή τι ενέργεια να κάνουμε.

Επίστρεψε ΑΥΣΤΗΡΑ ένα JSON αντικείμενο με την εξής δομή (χωρίς markdown):
{
  "voice_transcription": "Γράψε εδώ ΑΚΡΙΒΩΣ τι κατάλαβες ότι είπε το αφεντικό στον ήχο στα Ελληνικά",
  "summary": "Μια σύντομη σύνοψη της κατάστασης (Ποιος, γιατί και τι ζήτησε το αφεντικό)",
  "email_reply": "Η επίσημη, επαγγελματική απάντηση προς τον ${meta.senderName} λαμβάνοντας υπόψη όλο το ιστορικό συνομιλίας",
  "order_data": {
     "customer_name": "${meta.senderName}",
     "intent_why": "Ο λόγος της συνομιλίας με 3 λέξεις",
     "any_extra_form_data": "Οποιοδήποτε άλλο στοιχείο (ποσό, ημερομηνία) αναφέρθηκε στον ήχο"
  }
}`;

    const requestBody = {
        "contents": [
            { 
                "parts": [
                    { "text": prompt },
                    { "inlineData": { "mimeType": mimeType, "data": base64Audio } }
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
        const parsedData = JSON.parse(aiResultText);
        
        // 1. Εμφανίζουμε το Φωνητικό Input στο GUI
        document.getElementById("voiceInputText").innerText = parsedData.voice_transcription;
        document.getElementById("voiceInputContainer").style.display = "block";

        // 2. Εμφανίζουμε τα Δεδομένα της Φόρμας (για το ERP σας)
        document.getElementById("result").innerText = "Σύνοψη: " + parsedData.summary + "\n\nΔεδομένα Φόρμας: \n" + JSON.stringify(parsedData.order_data, null, 2);
        document.getElementById("formContainer").style.display = "block";

        // 3. Βάζουμε την Απάντηση στο Textarea
        document.getElementById("draftText").value = parsedData.email_reply;
        document.getElementById("draftContainer").style.display = "block";

        document.getElementById("status").innerText = "✨ Ανάλυση Ολοκληρώθηκε!";
        document.getElementById("status").style.color = "#32d74b";

    } catch (error) {
        document.getElementById("status").innerText = "Σφάλμα AI: " + error.message;
        document.getElementById("status").style.color = "#ff453a";
    }
}

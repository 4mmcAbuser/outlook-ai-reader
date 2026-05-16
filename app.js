// ==============================
// AI Voice Master - Outlook Add-in
// Full app.js
// ==============================

// Περιμένουμε να φορτώσει το Office
Office.onReady((info) => {

    if (info.host === Office.HostType.Outlook) {
        initApp();
    }

});

// ==============================
// GLOBALS
// ==============================

let apiKey = "";

const statusEl = document.getElementById("status");
const resultEl = document.getElementById("result");
const actionBtn = document.getElementById("actionBtn");
const saveKeyBtn = document.getElementById("saveKeyBtn");
const apiKeyInput = document.getElementById("apiKeyInput");

// Speech Recognition
let recognition = null;

// ==============================
// INIT
// ==============================

function initApp() {

    setStatus("Φόρτωση εφαρμογής...", "#0a84ff");

    // Load API key
    const savedKey = localStorage.getItem("geminiApiKey");

    if (savedKey) {

        apiKey = savedKey;
        apiKeyInput.value = savedKey;

        actionBtn.disabled = false;

        setStatus("Έτοιμο για χρήση!", "#32d74b");

    } else {

        actionBtn.disabled = true;

        setStatus("Πρόσθεσε Gemini API Key", "#ffcc00");
    }

    // Save API key
    saveKeyBtn.onclick = saveApiKey;

    // Init speech
    initSpeechRecognition();
}

// ==============================
// SAVE API KEY
// ==============================

function saveApiKey() {

    const key = apiKeyInput.value.trim();

    if (!key) {

        setStatus("Βάλε έγκυρο API Key", "#ff453a");
        return;
    }

    localStorage.setItem("geminiApiKey", key);

    apiKey = key;

    actionBtn.disabled = false;

    setStatus("Το API Key αποθηκεύτηκε!", "#32d74b");
}

// ==============================
// SPEECH RECOGNITION
// ==============================

function initSpeechRecognition() {

    const SpeechRecognitionAPI =
        window.SpeechRecognition ||
        window.webkitSpeechRecognition;

    if (!SpeechRecognitionAPI) {

        setStatus(
            "Το Outlook/WebView δεν υποστηρίζει Speech Recognition.",
            "#ff453a"
        );

        actionBtn.disabled = true;

        return;
    }

    recognition = new SpeechRecognitionAPI();

    recognition.lang = "el-GR";
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    // ==============================
    // BUTTON CLICK
    // ==============================

    actionBtn.onclick = async () => {

        if (!apiKey) {

            setStatus("Δεν υπάρχει API Key", "#ff453a");
            return;
        }

        try {

            // Ζήτα permission μικροφώνου
            await navigator.mediaDevices.getUserMedia({
                audio: true
            });

            resultEl.style.display = "none";

            setStatus("Ακούω... Μίλα τώρα!", "#ff453a");

            recognition.start();

        } catch (err) {

            console.error(err);

            setStatus(
                "Δεν επιτράπηκε η πρόσβαση στο μικρόφωνο.",
                "#ff453a"
            );
        }
    };

    // ==============================
    // START
    // ==============================

    recognition.onstart = () => {

        console.log("Speech recognition started");
    };

    // ==============================
    // RESULT
    // ==============================

    recognition.onresult = (event) => {

        try {

            const transcript =
                event.results[0][0].transcript;

            console.log("Voice Command:", transcript);

            setStatus(
                `Είπες: "${transcript}"`,
                "#0a84ff"
            );

            extractEmailAndProcess(transcript);

        } catch (err) {

            console.error(err);

            setStatus(
                "Σφάλμα αναγνώρισης φωνής.",
                "#ff453a"
            );
        }
    };

    // ==============================
    // ERROR
    // ==============================

    recognition.onerror = (event) => {

        console.error(event);

        let msg = "Speech Error";

        switch (event.error) {

            case "not-allowed":
                msg = "Δεν δόθηκε άδεια μικροφώνου.";
                break;

            case "network":
                msg = "Network error.";
                break;

            case "no-speech":
                msg = "Δεν ακούστηκε ομιλία.";
                break;

            case "audio-capture":
                msg = "Δεν βρέθηκε μικρόφωνο.";
                break;

            default:
                msg = "Speech Error: " + event.error;
        }

        setStatus(msg, "#ff453a");
    };

    // ==============================
    // END
    // ==============================

    recognition.onend = () => {

        console.log("Speech recognition ended");
    };
}

// ==============================
// READ EMAIL
// ==============================

function extractEmailAndProcess(voiceCommand) {

    setStatus(
        "Διαβάζω το email...",
        "#0a84ff"
    );

    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Text,

        function (asyncResult) {

            if (
                asyncResult.status ===
                Office.AsyncResultStatus.Succeeded
            ) {

                const emailBody = asyncResult.value;

                console.log("EMAIL:", emailBody);

                callGeminiAPI(
                    emailBody,
                    voiceCommand
                );

            } else {

                console.error(asyncResult.error);

                setStatus(
                    "Σφάλμα ανάγνωσης email.",
                    "#ff453a"
                );
            }
        }
    );
}

// ==============================
// GEMINI API
// ==============================

async function callGeminiAPI(
    emailText,
    voiceCommand
) {

    setStatus(
        "Στέλνω δεδομένα στο Gemini AI...",
        "#0a84ff"
    );

    const url =
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

    const prompt = `
Είσαι Executive Assistant.

EMAIL:
${emailText}

ΦΩΝΗΤΙΚΗ ΕΝΤΟΛΗ:
${voiceCommand}

Ανάλυσε το email και δημιούργησε:

1. summary
2. email_reply
3. action_items
4. order_data

Απάντησε ΜΟΝΟ σε JSON μορφή.
`;

    const requestBody = {

        contents: [
            {
                parts: [
                    {
                        text: prompt
                    }
                ]
            }
        ],

        generationConfig: {
            temperature: 0.4,
            responseMimeType: "application/json"
        }
    };

    try {

        const response = await fetch(url, {

            method: "POST",

            headers: {
                "Content-Type": "application/json"
            },

            body: JSON.stringify(requestBody)
        });

        const data = await response.json();

        console.log(data);

        if (data.error) {

            throw new Error(data.error.message);
        }

        const aiText =
            data.candidates[0]
                .content.parts[0].text;

        showResult(aiText);

        // Προαιρετικά:
        // αυτόματη εισαγωγή reply draft
        insertReplyIntoOutlook(aiText);

    } catch (err) {

        console.error(err);

        setStatus(
            "Σφάλμα AI: " + err.message,
            "#ff453a"
        );
    }
}

// ==============================
// SHOW RESULT
// ==============================

function showResult(text) {

    setStatus(
        "Το AI ολοκλήρωσε την επεξεργασία!",
        "#32d74b"
    );

    resultEl.style.display = "block";

    try {

        const parsed = JSON.parse(text);

        resultEl.innerText =
            JSON.stringify(parsed, null, 2);

    } catch {

        resultEl.innerText = text;
    }
}

// ==============================
// INSERT REPLY INTO OUTLOOK
// ==============================

function insertReplyIntoOutlook(aiText) {

    try {

        const parsed = JSON.parse(aiText);

        if (!parsed.email_reply) {
            return;
        }

        const replyText = parsed.email_reply;

        Office.context.mailbox.item.displayReplyForm({
            htmlBody:
                `
                <div style="font-family:Segoe UI;padding:10px;">
                    ${replyText.replace(/\n/g, "<br>")}
                </div>
                `
        });

    } catch (err) {

        console.error(
            "Reply insert error:",
            err
        );
    }
}

// ==============================
// STATUS HELPER
// ==============================

function setStatus(message, color = "#ffffff") {

    statusEl.innerText = message;
    statusEl.style.color = color;
}

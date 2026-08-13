/* Lao Dictate — Word Add-in
   - "Real-time" mode: TRUE streaming dictation using the Azure Speech SDK
     for JavaScript. It connects over a standard WebSocket (not a
     Chrome-only privileged API), so — unlike the browser's native
     webkitSpeechRecognition, which only authenticates from genuine Google
     Chrome and returns a generic "network" error everywhere else,
     including Word's own WebView2 — this works reliably inside the Word
     task pane on Windows, on Word on the web in any browser, and (with
     microphone permission) on Word for Mac too. Azure Speech officially
     lists Lao ("lo-LA") as a supported recognition locale. Requires an
     Azure Speech resource (key + region) — Azure offers a free tier.
   - "File upload" mode: sends a full audio file to OpenAI's transcription
     endpoint (works everywhere, but is not real-time).
*/

const LANG = "lo-LA"; // Azure Speech locale code for Lao
const WHISPER_LANG = "lo"; // OpenAI Whisper uses the shorter ISO-639-1 code

let recognizer = null;
let realtimeRunning = false;
let committedText = ""; // text already inserted into the Word document this session

Office.onReady(() => {
  console.log("[Lao Dictate] build 2 loaded — taskpane.js");

  document.getElementById("saveKeyBtn").onclick = saveKey;
  document.getElementById("saveAzureBtn").onclick = saveAzureConfig;
  document.getElementById("startRealtimeBtn").onclick = startRealtime;
  document.getElementById("stopRealtimeBtn").onclick = stopRealtime;
  document.getElementById("transcribeFileBtn").onclick = transcribeFile;

  const savedKey = localStorage.getItem("laoDictate_apiKey");
  if (savedKey) {
    document.getElementById("apiKey").value = savedKey;
    setKeyStatus("ใช้คีย์ที่บันทึกไว้");
  }

  const savedAzureKey = localStorage.getItem("laoDictate_azureKey");
  const savedAzureRegion = localStorage.getItem("laoDictate_azureRegion");
  if (savedAzureKey) document.getElementById("azureKey").value = savedAzureKey;
  if (savedAzureRegion) document.getElementById("azureRegion").value = savedAzureRegion;
  if (savedAzureKey && savedAzureRegion) {
    document.getElementById("azureStatus").textContent = "ใช้ค่าที่บันทึกไว้";
  }

  if (typeof SpeechSDK === "undefined") {
    document.getElementById("realtimeStatus").textContent =
      "โหลด Azure Speech SDK ไม่สำเร็จ (ตรวจสอบอินเทอร์เน็ต หรือว่า cdn.jsdelivr.net ถูกบล็อกหรือไม่)";
    document.getElementById("startRealtimeBtn").disabled = true;
  }
});

function saveAzureConfig() {
  const key = document.getElementById("azureKey").value.trim();
  const region = document.getElementById("azureRegion").value.trim();
  if (!key || !region) {
    document.getElementById("azureStatus").textContent = "กรุณากรอกทั้ง Key และ Region";
    return;
  }
  localStorage.setItem("laoDictate_azureKey", key);
  localStorage.setItem("laoDictate_azureRegion", region);
  document.getElementById("azureStatus").textContent = "บันทึกแล้ว ✓";
}

function saveKey() {
  const key = document.getElementById("apiKey").value.trim();
  if (!key) return;
  localStorage.setItem("laoDictate_apiKey", key);
  setKeyStatus("บันทึกแล้ว ✓");
}

function setKeyStatus(msg) {
  document.getElementById("keyStatus").textContent = msg;
}

function getKey() {
  const key = document.getElementById("apiKey").value.trim();
  if (!key) {
    alert("กรุณาใส่ OpenAI API Key ก่อน");
    throw new Error("missing API key");
  }
  return key;
}

/* ---------- Word insertion helper ---------- */
async function insertIntoDocument(text) {
  if (!text) return;
  const clean = text.trim();
  if (!clean) return;
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(clean + " ", Word.InsertLocation.end);
    context.document.getSelection().select(Word.SelectionMode.End);
    await context.sync();
  });
}

/* ---------- OpenAI transcription call ---------- */
async function transcribeBlob(blob, filename, apiKey) {
  const form = new FormData();
  form.append("file", blob, filename);
  form.append("model", "whisper-1");
  form.append("language", WHISPER_LANG);

  const resp = await fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "POST",
    headers: { Authorization: "Bearer " + apiKey },
    body: form,
  });

  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error("API error " + resp.status + ": " + errText);
  }
  const data = await resp.json();
  return data.text || "";
}

/* ---------- TRUE Real-time dictation (Azure Speech SDK) ---------- */
function startRealtime() {
  const statusEl = document.getElementById("realtimeStatus");
  const preview = document.getElementById("preview");

  const azureKey = localStorage.getItem("laoDictate_azureKey");
  const azureRegion = localStorage.getItem("laoDictate_azureRegion");

  if (!azureKey || !azureRegion) {
    statusEl.textContent = "กรุณาตั้งค่า Azure Speech Key และ Region ก่อน";
    return;
  }
  if (typeof SpeechSDK === "undefined") {
    statusEl.textContent = "Azure Speech SDK ยังโหลดไม่สำเร็จ";
    return;
  }

  committedText = "";
  preview.value = "";

  let speechConfig;
  try {
    speechConfig = SpeechSDK.SpeechConfig.fromSubscription(azureKey, azureRegion);
  } catch (e) {
    statusEl.textContent = "ตั้งค่า Azure Speech ไม่สำเร็จ: " + e.message;
    return;
  }
  speechConfig.speechRecognitionLanguage = LANG; // "lo-LA"

  const audioConfig = SpeechSDK.AudioConfig.fromDefaultMicrophoneInput();
  recognizer = new SpeechSDK.SpeechRecognizer(speechConfig, audioConfig);

  realtimeRunning = true;
  document.getElementById("startRealtimeBtn").disabled = true;
  document.getElementById("stopRealtimeBtn").disabled = false;
  statusEl.textContent = "🎙️ กำลังเชื่อมต่อ...";

  // Fires continuously while a phrase is still being recognized (live, word-by-word)
  recognizer.recognizing = (s, e) => {
    if (e.result.text) {
      preview.value = (committedText + " " + e.result.text).trim();
      statusEl.textContent = "🎙️ กำลังฟัง...";
    }
  };

  // Fires once a phrase/sentence boundary is finalized -> commit to the document
  recognizer.recognized = async (s, e) => {
    if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech && e.result.text) {
      committedText += (committedText ? " " : "") + e.result.text;
      preview.value = committedText;
      await insertIntoDocument(e.result.text);
    }
  };

  recognizer.canceled = (s, e) => {
    console.error("[Lao Dictate] recognizer canceled", e.reason, e.errorCode, e.errorDetails);
    statusEl.textContent =
      "เกิดข้อผิดพลาด (" + (e.errorCode || e.reason) + "): " + (e.errorDetails || "ไม่ทราบสาเหตุ");
    stopRealtime();
  };

  recognizer.sessionStarted = () => {
    statusEl.textContent = "🎙️ กำลังฟัง...";
  };

  recognizer.startContinuousRecognitionAsync(
    () => {
      /* started */
    },
    (err) => {
      statusEl.textContent = "เริ่มไม่สำเร็จ: " + err;
      stopRealtime();
    }
  );
}

function stopRealtime() {
  realtimeRunning = false;
  document.getElementById("startRealtimeBtn").disabled = false;
  document.getElementById("stopRealtimeBtn").disabled = true;
  document.getElementById("realtimeStatus").textContent = "หยุดแล้ว";

  if (recognizer) {
    recognizer.stopContinuousRecognitionAsync(
      () => {
        recognizer.close();
        recognizer = null;
      },
      () => {
        recognizer = null;
      }
    );
  }
}

/* ---------- File upload transcription ---------- */
async function transcribeFile() {
  const apiKey = getKey();
  const fileInput = document.getElementById("audioFile");
  const statusEl = document.getElementById("fileStatus");

  if (!fileInput.files || fileInput.files.length === 0) {
    statusEl.textContent = "กรุณาเลือกไฟล์เสียงก่อน";
    return;
  }

  const file = fileInput.files[0];
  // OpenAI transcription endpoint limit is 25 MB per request
  if (file.size > 25 * 1024 * 1024) {
    statusEl.textContent = "ไฟล์ใหญ่เกิน 25MB กรุณาตัดไฟล์เป็นช่วงสั้นลง";
    return;
  }

  statusEl.textContent = "กำลังถอดเสียง...";
  document.getElementById("transcribeFileBtn").disabled = true;

  try {
    const text = await transcribeBlob(file, file.name, apiKey);
    await insertIntoDocument(text);
    document.getElementById("preview").value = text.trim();
    statusEl.textContent = "แทรกข้อความลงเอกสารเรียบร้อย ✓";
  } catch (err) {
    statusEl.textContent = "เกิดข้อผิดพลาด: " + err.message;
  } finally {
    document.getElementById("transcribeFileBtn").disabled = false;
  }
}

/* Lao Dictate — Word Add-in
   - "Real-time" mode: TRUE streaming dictation using the browser's native
     Web Speech API (SpeechRecognition). Chrome/Edge route the audio to
     Google's speech recognition service, which supports Lao ("lo"). This
     gives genuine word-by-word interim results and sentence-by-sentence
     final results with no API key and no server needed. Only works where
     the task pane runs on a Chromium engine (Word desktop on Windows via
     WebView2, or Word on the web opened in Chrome/Edge). Safari-based
     WebViews (Word on Mac) do not implement this API — the file-upload
     mode below (OpenAI Whisper) is the fallback there.
   - "File upload" mode: sends a full audio file to OpenAI's transcription
     endpoint (works everywhere, but is not real-time).
*/

const LANG = "lo"; // used for both Web Speech API (lang) and Whisper (language)

let recognition = null;
let realtimeRunning = false;
let committedText = ""; // text already inserted into the Word document this session

Office.onReady(() => {
  document.getElementById("saveKeyBtn").onclick = saveKey;
  document.getElementById("startRealtimeBtn").onclick = startRealtime;
  document.getElementById("stopRealtimeBtn").onclick = stopRealtime;
  document.getElementById("transcribeFileBtn").onclick = transcribeFile;

  const savedKey = localStorage.getItem("laoDictate_apiKey");
  if (savedKey) {
    document.getElementById("apiKey").value = savedKey;
    setKeyStatus("ใช้คีย์ที่บันทึกไว้");
  }

  checkRealtimeSupport();
});

function checkRealtimeSupport() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  const statusEl = document.getElementById("realtimeStatus");
  if (!SR) {
    document.getElementById("startRealtimeBtn").disabled = true;
    statusEl.textContent =
      "เบราว์เซอร์/WebView นี้ไม่รองรับ Web Speech API (เช่น Word บน Mac) — กรุณาใช้โหมด 'ถอดเสียงจากไฟล์เสียง' แทน หรือเปิด Word บนเว็บด้วย Chrome/Edge";
  }
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
  form.append("language", LANG);

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

/* ---------- TRUE Real-time dictation (Web Speech API) ---------- */
function startRealtime() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  const statusEl = document.getElementById("realtimeStatus");
  const preview = document.getElementById("preview");

  if (!SR) {
    statusEl.textContent = "เบราว์เซอร์นี้ไม่รองรับ Web Speech API";
    return;
  }

  committedText = "";
  preview.value = "";

  recognition = new SR();
  recognition.lang = LANG; // "lo" — Lao
  recognition.continuous = true; // keep listening, don't stop after one phrase
  recognition.interimResults = true; // fire live word-by-word updates

  realtimeRunning = true;
  document.getElementById("startRealtimeBtn").disabled = true;
  document.getElementById("stopRealtimeBtn").disabled = false;
  statusEl.textContent = "🎙️ กำลังฟัง... พูดได้เลย";

  recognition.onresult = async (event) => {
    let interimText = "";

    for (let i = event.resultIndex; i < event.results.length; i++) {
      const result = event.results[i];
      const transcript = result[0].transcript;

      if (result.isFinal) {
        // A sentence/phrase is finalized -> commit it to the Word document immediately
        committedText += transcript;
        statusEl.textContent = "🎙️ กำลังฟัง...";
        await insertIntoDocument(transcript);
      } else {
        // Still being recognized -> show live, word-by-word, in the preview only
        interimText += transcript;
      }
    }

    // Live preview: committed (already in the doc) + what's currently being heard
    preview.value = (committedText + " " + interimText).trim();
  };

  recognition.onerror = (event) => {
    if (event.error === "no-speech") return; // benign, keep listening
    statusEl.textContent = "เกิดข้อผิดพลาด: " + event.error;
    if (event.error === "not-allowed" || event.error === "service-not-allowed") {
      stopRealtime();
    }
  };

  recognition.onend = () => {
    // Chrome auto-stops recognition periodically (e.g. after a pause) —
    // restart automatically to keep the session feeling continuous.
    if (realtimeRunning) {
      try {
        recognition.start();
      } catch (e) {
        /* already starting */
      }
    }
  };

  recognition.start();
}

function stopRealtime() {
  realtimeRunning = false;
  document.getElementById("startRealtimeBtn").disabled = false;
  document.getElementById("stopRealtimeBtn").disabled = true;
  document.getElementById("realtimeStatus").textContent = "หยุดแล้ว";

  if (recognition) {
    recognition.onend = null; // prevent auto-restart
    recognition.stop();
    recognition = null;
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

/**
 * Mail-auf-Desktop-Speichern – Outlook 365 Web-Add-in
 * taskpane.js
 *
 * Ablauf:
 *  1. Office.onReady → Mail-Infos & Anhänge laden
 *  2. Benutzer wählt Speichermodus
 *  3. Dateien werden via File System Access API (showDirectoryPicker)
 *     oder als Download-Blob auf den Desktop gespeichert
 *  4. Mail wird nach erfolgreichem Speichern in Outlook gelöscht
 */

/* ─────────────────────────────────────────────
   STATE
───────────────────────────────────────────── */
const state = {
  item: null,           // Office.context.mailbox.item
  subject: "",
  from: "",
  date: "",
  body: "",             // EML-Body (Text)
  bodyHtml: "",
  attachments: [],      // [{ id, name, size, contentType, content(base64) }]
  mode: null,           // "mail-only" | "attachments-only" | "both" | "select"
  selectedAttachments: new Set(),
  includeMailInSelect: false,
};

/* ─────────────────────────────────────────────
   DOM HELPERS
───────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);
function showScreen(name) {
  document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
  $(name).classList.add("active");
}
function show(id)  { $(id).classList.remove("hidden"); }
function hide(id)  { $(id).classList.add("hidden"); }
function text(id, v) { $(id).textContent = v; }

/* ─────────────────────────────────────────────
   OFFICE INIT
───────────────────────────────────────────── */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initAddin();
  }
});

async function initAddin() {
  showScreen("screen-loading");

  try {
    state.item = Office.context.mailbox.item;

    // Subject
    state.subject = await getAsync(state.item.subject);

    // From
    const from = await getAsync(state.item.from);
    state.from = from?.emailAddress || from?.displayName || "Unbekannt";

    // Date
    const dateReceived = state.item.dateTimeCreated;
    state.date = dateReceived
      ? new Date(dateReceived).toLocaleString("de-DE")
      : "–";

    // Body (HTML für EML)
    state.bodyHtml = await getBodyAsync("html");
    state.body = await getBodyAsync("text");

    // Attachments metadata
    state.attachments = (state.item.attachments || [])
      .filter(a => !a.isInline)
      .map(a => ({ id: a.id, name: a.name, size: a.size, contentType: a.contentType, content: null }));

    renderMain();
    showScreen("screen-main");
  } catch (err) {
    showError("Fehler beim Laden der Mail: " + err.message);
  }
}

/* Promisify Office callbacks */
function getAsync(prop) {
  return new Promise((res, rej) => {
    if (typeof prop?.getAsync === "function") {
      prop.getAsync((r) => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(new Error(r.error?.message)));
    } else {
      res(prop); // already a value
    }
  });
}

function getBodyAsync(type) {
  const coercionType = type === "html" ? Office.CoercionType.Html : Office.CoercionType.Text;
  return new Promise((res, rej) => {
    state.item.body.getAsync(coercionType, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? res(r.value || "") : rej(new Error(r.error?.message))
    );
  });
}

/* ─────────────────────────────────────────────
   RENDER MAIN SCREEN
───────────────────────────────────────────── */
function renderMain() {
  text("mail-subject", state.subject || "(Kein Betreff)");
  text("mail-from", state.from);
  text("mail-date", state.date);

  if (state.attachments.length === 0) {
    show("no-attachments-block");
    hide("attachments-block");
  } else {
    hide("no-attachments-block");
    show("attachments-block");
    renderAttachmentList();
  }
  bindEvents();
}

function renderAttachmentList() {
  const list = $("attachment-list");
  list.innerHTML = "";
  state.attachments.forEach((att) => {
    const item = document.createElement("label");
    item.className = "attachment-item";
    item.dataset.id = att.id;

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = att.id;
    cb.addEventListener("change", onAttachmentToggle);

    const nameSpan = document.createElement("span");
    nameSpan.className = "att-name";
    nameSpan.title = att.name;
    nameSpan.textContent = att.name;

    const sizeSpan = document.createElement("span");
    sizeSpan.className = "att-size";
    sizeSpan.textContent = formatSize(att.size);

    item.append(cb, nameSpan, sizeSpan);
    list.appendChild(item);
  });
}

function onAttachmentToggle(e) {
  const id = e.target.value;
  const label = e.target.closest(".attachment-item");
  if (e.target.checked) {
    state.selectedAttachments.add(id);
    label.classList.add("checked");
  } else {
    state.selectedAttachments.delete(id);
    label.classList.remove("checked");
  }
  updateSaveSelectedBtn();
}

function updateSaveSelectedBtn() {
  const hasSelection = state.selectedAttachments.size > 0 || $("include-mail-check").checked;
  $("btn-save-selected").disabled = !hasSelection;
}

/* ─────────────────────────────────────────────
   EVENT BINDING
───────────────────────────────────────────── */
function bindEvents() {
  // No-attachment path
  $("btn-save-mail-only").addEventListener("click", () => execute("mail-only"));

  // Choice cards
  const cards = {
    "choice-mail-only": "mail-only",
    "choice-attachments-only": "attachments-only",
    "choice-both": "both",
    "choice-select": "select",
  };

  Object.entries(cards).forEach(([id, mode]) => {
    $(id).addEventListener("click", () => selectMode(mode));
  });

  $("include-mail-check").addEventListener("change", updateSaveSelectedBtn);
  $("btn-save-selected").addEventListener("click", executeSelect);
  $("btn-execute").addEventListener("click", () => execute(state.mode));

  $("btn-done").addEventListener("click", () => window.close());
  $("btn-retry").addEventListener("click", () => {
    state.mode = null;
    state.selectedAttachments.clear();
    showScreen("screen-main");
  });
}

function selectMode(mode) {
  state.mode = mode;

  // Highlight active card
  document.querySelectorAll(".choice-card").forEach(c => c.classList.remove("active"));
  const idMap = { "mail-only": "choice-mail-only", "attachments-only": "choice-attachments-only", "both": "choice-both", "select": "choice-select" };
  $(idMap[mode])?.classList.add("active");

  if (mode === "select") {
    show("attachment-selector");
    hide("action-buttons");
  } else {
    hide("attachment-selector");
    const labels = { "mail-only": "Mail speichern (.eml)", "attachments-only": "Alle Anhänge speichern", "both": "Mail + Anhänge speichern" };
    text("btn-execute-label", labels[mode] || "Speichern");
    show("action-buttons");
  }
}

/* ─────────────────────────────────────────────
   EXECUTE SAVE
───────────────────────────────────────────── */
async function execute(mode) {
  showScreen("screen-progress");
  text("progress-text", "Wird vorbereitet …");

  try {
    const savedFiles = [];

    if (mode === "mail-only" || mode === "both") {
      text("progress-text", "Mail wird erstellt …");
      const emlBlob = buildEml();
      const filename = sanitizeFilename(state.subject || "Mail") + ".eml";
      await saveFile(emlBlob, filename);
      savedFiles.push(filename);
    }

    if (mode === "attachments-only" || mode === "both") {
      const atts = state.attachments;
      for (let i = 0; i < atts.length; i++) {
        text("progress-text", `Anhang ${i + 1}/${atts.length}: ${atts[i].name}`);
        const content = await loadAttachment(atts[i].id);
        const blob = base64ToBlob(content, atts[i].contentType);
        await saveFile(blob, atts[i].name);
        savedFiles.push(atts[i].name);
      }
    }

    await deleteMail();
    showSuccess(savedFiles);
  } catch (err) {
    showError(err.message);
  }
}

async function executeSelect() {
  showScreen("screen-progress");
  text("progress-text", "Wird vorbereitet …");

  try {
    const savedFiles = [];

    if ($("include-mail-check").checked) {
      text("progress-text", "Mail wird erstellt …");
      const emlBlob = buildEml();
      const filename = sanitizeFilename(state.subject || "Mail") + ".eml";
      await saveFile(emlBlob, filename);
      savedFiles.push(filename);
    }

    const selected = [...state.selectedAttachments];
    for (let i = 0; i < selected.length; i++) {
      const att = state.attachments.find(a => a.id === selected[i]);
      if (!att) continue;
      text("progress-text", `Anhang ${i + 1}/${selected.length}: ${att.name}`);
      const content = await loadAttachment(att.id);
      const blob = base64ToBlob(content, att.contentType);
      await saveFile(blob, att.name);
      savedFiles.push(att.name);
    }

    await deleteMail();
    showSuccess(savedFiles);
  } catch (err) {
    showError(err.message);
  }
}

/* ─────────────────────────────────────────────
   LOAD ATTACHMENT CONTENT
───────────────────────────────────────────── */
function loadAttachment(attachmentId) {
  return new Promise((res, rej) => {
    state.item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        res(result.value.content); // base64
      } else {
        rej(new Error("Anhang konnte nicht geladen werden: " + result.error?.message));
      }
    });
  });
}

/* ─────────────────────────────────────────────
   BUILD EML
───────────────────────────────────────────── */
function buildEml() {
  const boundary = "----=_Part_" + Date.now();
  const dateStr = state.item.dateTimeCreated
    ? new Date(state.item.dateTimeCreated).toUTCString()
    : new Date().toUTCString();

  let eml = [
    `From: ${state.from}`,
    `To: ${Office.context.mailbox.userProfile?.emailAddress || ""}`,
    `Subject: ${state.subject || "(Kein Betreff)"}`,
    `Date: ${dateStr}`,
    `MIME-Version: 1.0`,
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    ``,
    `--${boundary}`,
    `Content-Type: text/html; charset=UTF-8`,
    `Content-Transfer-Encoding: quoted-printable`,
    ``,
    state.bodyHtml || state.body || "(Kein Inhalt)",
    ``,
  ].join("\r\n");

  eml += `\r\n--${boundary}--\r\n`;
  return new Blob([eml], { type: "message/rfc822" });
}

/* ─────────────────────────────────────────────
   FILE SAVING
   Tries File System Access API (Desktop) first,
   falls back to download-trigger approach.
───────────────────────────────────────────── */
async function saveFile(blob, filename) {
  // Try modern File System Access API (Chrome/Edge with user gesture)
  if (window.showSaveFilePicker) {
    try {
      const ext = filename.split(".").pop();
      const mimeMap = { eml: "message/rfc822", pdf: "application/pdf" };
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        startIn: "desktop",
        types: [{ description: "Datei", accept: { [mimeMap[ext] || "application/octet-stream"]: ["." + ext] } }],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return;
    } catch (err) {
      if (err.name === "AbortError") throw new Error("Speichern abgebrochen.");
      // Fall through to download fallback
    }
  }

  // Fallback: trigger browser download
  downloadBlob(blob, filename);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

/* ─────────────────────────────────────────────
   DELETE MAIL
───────────────────────────────────────────── */
function deleteMail() {
  return new Promise((res, rej) => {
    // moveToFolder to Deleted Items via REST, or use Office JS
    // Office.js 1.5+: item.moveAsync (if available)
    if (state.item.moveAsync) {
      state.item.moveAsync(Office.MailboxEnums.MoveOperationType.MoveToDeletedItems, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          res();
        } else {
          // Non-fatal: warn but don't block success
          console.warn("Mail konnte nicht gelöscht werden:", result.error?.message);
          res();
        }
      });
    } else {
      // Fallback: EWS delete via REST (requires additional token)
      // For now, resolve and show note to user
      console.warn("moveAsync nicht verfügbar – Mail wurde NICHT automatisch gelöscht.");
      res();
    }
  });
}

/* ─────────────────────────────────────────────
   SUCCESS / ERROR
───────────────────────────────────────────── */
function showSuccess(files) {
  const list = $("saved-files-list");
  list.innerHTML = "";
  files.forEach(f => {
    const d = document.createElement("div");
    d.className = "saved-file-entry";
    d.textContent = f;
    list.appendChild(d);
  });
  showScreen("screen-success");
}

function showError(msg) {
  text("error-message", msg);
  showScreen("screen-error");
}

/* ─────────────────────────────────────────────
   HELPERS
───────────────────────────────────────────── */
function base64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: mimeType || "application/octet-stream" });
}

function sanitizeFilename(name) {
  return name.replace(/[\\/:*?"<>|]/g, "_").substring(0, 120);
}

function formatSize(bytes) {
  if (!bytes) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

/**
 * HERO Mail Sync – taskpane.js
 * GraphQL-basierte Projektsuche + PDF-Upload
 */

const HERO_GQL = "/api/hero";
const STORAGE_KEY = "hero_addin_apikey";

let apiKey = "";
let uploadMutationName = null; // wird per Introspection ermittelt
let selectedProject = null;
let mailData = {};
let searchTimer = null;

// ─── INIT ────────────────────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadSettings();
    loadMailData();
  }
});

// Fallback für lokale Tests ohne Outlook
if (typeof Office === "undefined") {
  document.addEventListener("DOMContentLoaded", () => {
    loadSettings();
    mailData = {
      subject: "[Test] Anfrage Elektroinstallation Neubau",
      from: "max.mustermann@example.de",
      fromName: "Max Mustermann",
      date: new Date().toLocaleString("de-DE"),
      body: "Sehr geehrte Damen und Herren,\n\nanbei die Unterlagen für das Projekt.",
      attachments: []
    };
    renderMailPreview();
  });
}

// ─── SETTINGS ────────────────────────────────────────────────────────────────

function loadSettings() {
  try {
    apiKey = localStorage.getItem(STORAGE_KEY) || "";
    document.getElementById("apiKeyInput").value = apiKey;
  } catch (e) {}

  updateConnectionUI(apiKey ? "connected" : "disconnected");

  if (!apiKey) {
    toggleSettings();
  } else {
    // Verbindung im Hintergrund testen + Upload-Mutation ermitteln
    testConnectionAndIntrospect();
  }
}

function saveSettings() {
  apiKey = document.getElementById("apiKeyInput").value.trim();
  try { localStorage.setItem(STORAGE_KEY, apiKey); } catch (e) {}

  updateConnectionUI("checking");
  testConnectionAndIntrospect().then((ok) => {
    if (ok) {
      showStatus("success", "✅ Verbindung zu HERO erfolgreich!");
      toggleSettings();
    } else {
      showStatus("error", "❌ Verbindung fehlgeschlagen – API-Key prüfen.");
    }
  });
}

function toggleSettings() {
  document.getElementById("settingsPanel").classList.toggle("open");
}

async function testConnectionAndIntrospect() {
  try {
    // Verbindungstest
    const test = await heroQuery(`{ contacts(first: 1) { id } }`);
    if (!test.data) {
      updateConnectionUI("disconnected");
      return false;
    }
    updateConnectionUI("connected");

    // Upload-Mutation per Introspection ermitteln (einmalig)
    if (!uploadMutationName) {
      uploadMutationName = await findUploadMutation();
    }
    return true;
  } catch (e) {
    updateConnectionUI("disconnected");
    return false;
  }
}

/**
 * Durchsucht das GraphQL-Schema nach einer File-Upload-Mutation.
 * Gibt den Mutations-Namen zurück oder null falls nicht gefunden.
 */
async function findUploadMutation() {
  try {
    const result = await heroQuery(`
      {
        __schema {
          mutationType {
            fields {
              name
              args { name type { name kind ofType { name kind } } }
            }
          }
        }
      }
    `);

    const fields = result?.data?.__schema?.mutationType?.fields || [];
    const keywords = ["upload", "file", "document", "attachment", "pdf"];

    const match = fields.find((f) =>
      keywords.some((kw) => f.name.toLowerCase().includes(kw))
    );

    if (match) {
      console.log("HERO Upload-Mutation gefunden:", match.name, match.args.map(a => a.name));
      return match.name;
    }
  } catch (e) {
    console.warn("Introspection fehlgeschlagen:", e);
  }
  return null;
}

function updateConnectionUI(state) {
  const dot = document.getElementById("connDot");
  const text = document.getElementById("connText");
  dot.className = "dot " + state;
  text.textContent = state === "connected" ? "Verbunden" : state === "checking" ? "Prüfe..." : "Nicht verbunden";
}

// ─── MAIL DATA ────────────────────────────────────────────────────────────────

function loadMailData() {
  try {
    const item = Office.context.mailbox.item;

    mailData.subject = item.subject || "(Kein Betreff)";
    mailData.from = item.from?.emailAddress || "";
    mailData.fromName = item.from?.displayName || "";
    mailData.date = item.dateTimeCreated
      ? item.dateTimeCreated.toLocaleString("de-DE")
      : "";

    mailData.attachments = [];
    if (item.attachments) {
      for (const att of item.attachments) {
        if (!att.isInline) {
          mailData.attachments.push({
            id: att.id,
            name: att.name,
            size: att.size,
            contentType: att.contentType
          });
        }
      }
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      mailData.body = result.status === Office.AsyncResultStatus.Succeeded
        ? result.value
        : "";
      renderMailPreview();
    });
  } catch (e) {
    mailData.subject = "Fehler beim Laden";
    renderMailPreview();
  }
}

function renderMailPreview() {
  document.getElementById("emailSubject").textContent = mailData.subject || "";
  document.getElementById("emailFrom").textContent =
    mailData.fromName ? `${mailData.fromName} <${mailData.from}>` : (mailData.from || "–");
  document.getElementById("emailDate").textContent = mailData.date || "–";

  const row = document.getElementById("attachmentsRow");
  if (mailData.attachments && mailData.attachments.length > 0) {
    row.style.display = "flex";
    row.innerHTML = mailData.attachments.map((a) =>
      `<span class="att-badge">📎 ${escHtml(a.name)} (${fmtSize(a.size)})</span>`
    ).join("");
  }
}

// ─── GRAPHQL ──────────────────────────────────────────────────────────────────

async function heroQuery(query, variables) {
  const body = variables
    ? JSON.stringify({ query, variables })
    : JSON.stringify({ query });

  const res = await fetch(HERO_GQL, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body
  });
  return res.json();
}

/**
 * Führt eine Query aus und gibt das Daten-Array unter `dataKey` zurück.
 * Gibt null zurück wenn die Query fehlschlägt oder leer ist.
 */
async function tryQuery(query, variables, dataKey) {
  try {
    const result = await heroQuery(query, variables);
    if (result.errors) {
      console.warn(`Query "${dataKey}" Fehler:`, result.errors);
      return null;
    }
    const data = result.data?.[dataKey];
    console.log(`Query "${dataKey}" → ${data?.length ?? "?"} Einträge`);
    return Array.isArray(data) ? data : null;
  } catch (e) {
    console.warn(`Query "${dataKey}" Exception:`, e);
    return null;
  }
}

function matchesSearch(p, s) {
  const nr = (p.project_nr || "").toLowerCase();
  const company = (p.customer?.company_name || "").toLowerCase();
  const name = (`${p.customer?.first_name || ""} ${p.customer?.last_name || ""}`).toLowerCase();
  const city = (p.address?.city || "").toLowerCase();
  return nr.includes(s) || company.includes(s) || name.includes(s) || city.includes(s);
}

// ─── PROJECT SEARCH ──────────────────────────────────────────────────────────

function handleSearch() {
  const term = document.getElementById("searchInput").value.trim();
  if (searchTimer) clearTimeout(searchTimer);

  if (term.length < 2) {
    document.getElementById("projectResults").innerHTML = "";
    document.getElementById("searchHint").textContent = "Mindestens 2 Zeichen eingeben";
    return;
  }

  document.getElementById("searchHint").textContent = "";
  document.getElementById("projectResults").innerHTML =
    `<div class="spinner-wrap"><div class="spinner"></div>Suche läuft...</div>`;

  searchTimer = setTimeout(() => searchProjects(term), 350);
}

async function searchProjects(term) {
  if (!apiKey) {
    showStatus("error", "Bitte zuerst API-Key in den Einstellungen hinterlegen.");
    return;
  }

  const s = term.toLowerCase();

  try {
    // Versuch 1: project_matches mit search-Parameter
    let projects = await tryQuery(`
      query ($search: String) {
        project_matches(search: $search, first: 50) {
          id project_nr
          customer { first_name last_name company_name }
          address { city }
          current_project_match_status { name }
        }
      }`, { search: term }, "project_matches");

    // Versuch 2: project_matches ohne search (clientseitig filtern)
    if (!projects) {
      const all = await tryQuery(`{
        project_matches(first: 500) {
          id project_nr
          customer { first_name last_name company_name }
          address { city }
          current_project_match_status { name }
        }
      }`, null, "project_matches");

      projects = (all || []).filter((p) => matchesSearch(p, s));
    }

    // Versuch 3: alternatives Query-Feld "projects"
    if (!projects) {
      const all = await tryQuery(`{
        projects(first: 500) {
          id project_nr
          customer { first_name last_name company_name }
          address { city }
          current_project_match_status { name }
        }
      }`, null, "projects");

      projects = (all || []).filter((p) => matchesSearch(p, s));
    }

    if (!projects) {
      document.getElementById("projectResults").innerHTML =
        `<div class="state-msg">Keine Ergebnisse. API-Antwort in der Browser-Konsole prüfen.</div>`;
      return;
    }

    renderProjects(projects);
  } catch (e) {
    console.error("searchProjects Fehler:", e);
    document.getElementById("projectResults").innerHTML =
      `<div class="state-msg">Fehler: ${e.message}</div>`;
  }
}

function renderProjects(projects) {
  const container = document.getElementById("projectResults");

  if (!projects.length) {
    container.innerHTML = `<div class="state-msg">Keine Projekte gefunden</div>`;
    return;
  }

  container.innerHTML = projects.map((p) => {
    const customerName = p.customer
      ? (p.customer.company_name ||
         `${p.customer.first_name || ""} ${p.customer.last_name || ""}`.trim())
      : "Unbekannt";
    const status = p.current_project_match_status?.name || "";
    const sel = selectedProject?.id === p.id ? "selected" : "";

    return `
      <div class="project-item ${sel}"
           onclick="selectProject(${p.id}, '${escAttr(p.project_nr || "")}', '${escAttr(customerName)}')">
        <span class="project-nr">${escHtml(p.project_nr || "—")}</span>
        <div class="project-info">
          <div class="project-name">${escHtml(customerName)}</div>
          <div class="project-status">${escHtml(status)}</div>
        </div>
        <div class="project-check"></div>
      </div>`;
  }).join("");
}

function selectProject(id, nr, name) {
  selectedProject = { id, nr, name };

  // Visuell markieren
  document.querySelectorAll(".project-item").forEach((el) => el.classList.remove("selected"));
  event.currentTarget.classList.add("selected");

  const btn = document.getElementById("submitBtn");
  btn.disabled = false;
  btn.textContent = `📤 An ${nr} senden`;
}

// ─── PDF GENERATION ───────────────────────────────────────────────────────────

/**
 * Erstellt ein PDF aus den Mail-Daten und gibt es als Base64-String zurück.
 */
function generateEmailPdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });

  const pageW = doc.internal.pageSize.getWidth();
  const margin = 15;
  const maxW = pageW - margin * 2;
  let y = 20;

  // Titel
  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.setTextColor(245, 155, 0);
  doc.text("E-Mail", margin, y);
  y += 8;

  // Trennlinie
  doc.setDrawColor(226, 229, 234);
  doc.line(margin, y, pageW - margin, y);
  y += 6;

  // Header-Felder
  doc.setFontSize(10);
  doc.setTextColor(95, 102, 114);

  const fields = [
    ["Von:", `${mailData.fromName || ""} <${mailData.from || ""}>`],
    ["Betreff:", mailData.subject || ""],
    ["Datum:", mailData.date || ""],
  ];

  for (const [label, value] of fields) {
    doc.setFont("helvetica", "bold");
    doc.text(label, margin, y);
    doc.setFont("helvetica", "normal");
    const lines = doc.splitTextToSize(value, maxW - 25);
    doc.text(lines, margin + 22, y);
    y += lines.length * 5 + 2;
  }

  if (mailData.attachments && mailData.attachments.length > 0) {
    doc.setFont("helvetica", "bold");
    doc.text("Anhänge:", margin, y);
    doc.setFont("helvetica", "normal");
    const attNames = mailData.attachments.map((a) => a.name).join(", ");
    const lines = doc.splitTextToSize(attNames, maxW - 25);
    doc.text(lines, margin + 22, y);
    y += lines.length * 5 + 2;
  }

  y += 4;
  doc.setDrawColor(226, 229, 234);
  doc.line(margin, y, pageW - margin, y);
  y += 7;

  // Body
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(26, 29, 35);

  const bodyText = (mailData.body || "").substring(0, 8000);
  const bodyLines = doc.splitTextToSize(bodyText, maxW);

  for (const line of bodyLines) {
    if (y > 275) {
      doc.addPage();
      y = 20;
    }
    doc.text(line, margin, y);
    y += 4.5;
  }

  return doc.output("datauristring").split(",")[1]; // Base64
}

// ─── SUBMIT ───────────────────────────────────────────────────────────────────

async function submitToHero() {
  if (!selectedProject || !apiKey) return;

  const doLogbook = document.getElementById("optLogbook").checked;
  const doPdf = document.getElementById("optPdf").checked;
  const doAttachments = document.getElementById("optAttachments").checked;

  const btn = document.getElementById("submitBtn");
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner" style="width:14px;height:14px;border-width:2px;border-top-color:white"></span> Wird gesendet...`;

  try {
    // 1. Logbuch-Eintrag
    if (doLogbook) {
      showStatus("loading", "Logbuch-Eintrag wird erstellt...");
      await createLogbookEntry();
    }

    // 2. E-Mail als PDF
    if (doPdf) {
      showStatus("loading", "E-Mail als PDF wird generiert und hochgeladen...");
      await uploadEmailAsPdf();
    }

    // 3. Anhänge
    if (doAttachments && mailData.attachments && mailData.attachments.length > 0) {
      for (let i = 0; i < mailData.attachments.length; i++) {
        showStatus("loading", `Anhang ${i + 1}/${mailData.attachments.length} wird hochgeladen...`);
        try {
          await uploadAttachment(mailData.attachments[i]);
        } catch (e) {
          console.warn("Anhang-Upload übersprungen:", mailData.attachments[i].name, e);
        }
      }
    }

    showStatus("success", `✅ E-Mail wurde Projekt ${selectedProject.nr} zugeordnet!`);
    btn.textContent = `📤 An ${selectedProject.nr} senden`;
    btn.disabled = false;

  } catch (e) {
    showStatus("error", "❌ Fehler: " + e.message);
    btn.textContent = `📤 An ${selectedProject.nr} senden`;
    btn.disabled = false;
  }
}

// Gecachte Argument-Namen für add_logbook_entry
let logbookArgs = null;

async function resolveLogbookArgs() {
  if (logbookArgs) return logbookArgs;

  const result = await heroQuery(`{
    __schema {
      mutationType {
        fields {
          name
          args { name type { name kind ofType { name } } }
        }
      }
    }
  }`);

  const fields = result?.data?.__schema?.mutationType?.fields || [];
  const mut = fields.find((f) => f.name === "add_logbook_entry");

  if (!mut) throw new Error("Mutation add_logbook_entry nicht im Schema gefunden.");

  console.log("add_logbook_entry args:", mut.args.map((a) => a.name));

  // Argument-Namen flexibel erkennen
  const find = (keywords) =>
    mut.args.find((a) => keywords.some((kw) => a.name.toLowerCase().includes(kw)))?.name;

  logbookArgs = {
    projectId: find(["project_match_id", "project_id", "projectmatch", "project"]),
    message:   find(["message", "text", "content", "body", "note"]),
  };

  if (!logbookArgs.projectId || !logbookArgs.message) {
    throw new Error(
      `Konnte Argument-Namen nicht zuordnen. Verfügbar: ${mut.args.map((a) => a.name).join(", ")}`
    );
  }

  return logbookArgs;
}

async function createLogbookEntry() {
  const parts = [
    "📧 E-Mail zugeordnet",
    "─".repeat(40),
    `Von: ${mailData.fromName || ""} <${mailData.from || ""}>`,
    `Betreff: ${mailData.subject || ""}`,
    `Datum: ${mailData.date || ""}`,
  ];

  if (mailData.attachments && mailData.attachments.length > 0) {
    parts.push(`Anhänge: ${mailData.attachments.map((a) => a.name).join(", ")}`);
  }

  parts.push("─".repeat(40));

  if (mailData.body) {
    const body = mailData.body.length > 3000
      ? mailData.body.substring(0, 3000) + "\n...(gekürzt)"
      : mailData.body;
    parts.push(body);
  }

  const args = await resolveLogbookArgs();
  const message = parts.join("\n");

  const result = await heroQuery(`
    mutation ($pid: Int!, $msg: String!) {
      add_logbook_entry(${args.projectId}: $pid, ${args.message}: $msg) { id }
    }
  `, { pid: selectedProject.id, msg: message });

  if (result.errors) {
    throw new Error("Logbuch-Fehler: " + JSON.stringify(result.errors));
  }
}

async function uploadEmailAsPdf() {
  const pdfBase64 = generateEmailPdf();
  const filename = `Email_${sanitizeFilename(mailData.subject)}_${new Date().toISOString().slice(0,10)}.pdf`;
  await uploadFileToHero(filename, pdfBase64, "application/pdf");
}

async function uploadAttachment(attachment) {
  return new Promise((resolve, reject) => {
    if (typeof Office !== "undefined" && Office.context?.mailbox) {
      Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          uploadFileToHero(attachment.name, result.value.content, attachment.contentType || "application/octet-stream")
            .then(resolve)
            .catch(reject);
        } else {
          reject(new Error("Konnte Anhang nicht laden"));
        }
      });
    } else {
      resolve(); // Test-Modus
    }
  });
}

/**
 * Lädt eine Datei (Base64) via HERO GraphQL hoch.
 * Versucht alle bekannten Mutations-Namen und nutzt ggf. den per Introspection gefundenen.
 */
async function uploadFileToHero(filename, base64Content, contentType) {
  // Kandidaten in absteigender Wahrscheinlichkeit
  const candidates = [
    uploadMutationName,
    "upload_project_file",
    "upload_document",
    "add_document_to_project",
    "add_file_to_project_match",
    "upload_file",
  ].filter(Boolean);

  for (const mutName of candidates) {
    try {
      const result = await heroQuery(`
        mutation ($projectMatchId: Int!, $filename: String!, $content: String!, $contentType: String!) {
          ${mutName}(
            project_match_id: $projectMatchId,
            filename: $filename,
            content: $content,
            content_type: $contentType
          ) { id }
        }
      `, {
        projectMatchId: selectedProject.id,
        filename,
        content: base64Content,
        contentType
      });

      if (!result.errors) {
        console.log(`Upload erfolgreich via Mutation: ${mutName}`);
        uploadMutationName = mutName; // Für nächste Uploads merken
        return result;
      }
    } catch (e) {
      // nächsten Kandidaten versuchen
    }
  }

  // Alle Versuche fehlgeschlagen – als Logbuch-Notiz vermerken
  console.warn("Kein gültiger Upload-Endpoint gefunden. Bitte HERO Support kontaktieren.");
  await heroQuery(`
    mutation ($projectMatchId: Int!, $message: String!) {
      add_logbook_entry(project_match_id: $projectMatchId, message: $message) { id }
    }
  `, {
    projectMatchId: selectedProject.id,
    message: `📎 Datei konnte nicht hochgeladen werden: ${filename}\nBitte prüfen Sie die API-Dokumentation für den Upload-Endpoint.`
  });
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function showStatus(type, message) {
  const bar = document.getElementById("statusBar");
  bar.className = `status-bar show ${type}`;

  const icon = type === "loading"
    ? `<div class="spinner" style="width:14px;height:14px;flex-shrink:0"></div>`
    : "";
  bar.innerHTML = icon + message;

  if (type !== "loading") {
    setTimeout(() => bar.classList.remove("show"), 6000);
  }
}

function fmtSize(bytes) {
  if (!bytes) return "?";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}

function escHtml(str) {
  const d = document.createElement("div");
  d.textContent = str;
  return d.innerHTML;
}

function escAttr(str) {
  return String(str).replace(/'/g, "\\'").replace(/"/g, "&quot;");
}

function sanitizeFilename(str) {
  return (str || "Mail").replace(/[^a-zA-Z0-9äöüÄÖÜß _-]/g, "_").substring(0, 60);
}

/**
 * HERO Mail Sync – taskpane.js
 */

const HERO_GQL = "/api/hero";
const STORAGE_KEY = "hero_addin_apikey";

let apiKey = "";
let selectedProject = null;
let mailData = {};
let searchTimer = null;

// ─── INIT ─────────────────────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadSettings();
    loadMailData();
  }
});

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

// ─── SETTINGS ─────────────────────────────────────────────────────────────────

function loadSettings() {
  try {
    apiKey = localStorage.getItem(STORAGE_KEY) || "";
    document.getElementById("apiKeyInput").value = apiKey;
  } catch (e) {}

  updateConnectionUI(apiKey ? "connected" : "disconnected");
  if (!apiKey) toggleSettings();
  else testConnection();
}

function saveSettings() {
  apiKey = document.getElementById("apiKeyInput").value.trim();
  try { localStorage.setItem(STORAGE_KEY, apiKey); } catch (e) {}

  updateConnectionUI("checking");
  testConnection().then((ok) => {
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

async function testConnection() {
  try {
    const test = await heroQuery(`{ contacts(first: 1) { id } }`);
    const ok = !!test.data;
    updateConnectionUI(ok ? "connected" : "disconnected");
    return ok;
  } catch (e) {
    updateConnectionUI("disconnected");
    return false;
  }
}

function updateConnectionUI(state) {
  document.getElementById("connDot").className = "dot " + state;
  document.getElementById("connText").textContent =
    state === "connected" ? "Verbunden" :
    state === "checking"  ? "Prüfe..." : "Nicht verbunden";
}

// ─── MAIL DATA ────────────────────────────────────────────────────────────────

function loadMailData() {
  try {
    const item = Office.context.mailbox.item;
    mailData.subject  = item.subject || "(Kein Betreff)";
    mailData.from     = item.from?.emailAddress || "";
    mailData.fromName = item.from?.displayName || "";
    mailData.date     = item.dateTimeCreated
      ? item.dateTimeCreated.toLocaleString("de-DE") : "";

    mailData.attachments = [];
    if (item.attachments) {
      for (const att of item.attachments) {
        if (!att.isInline) {
          mailData.attachments.push({
            id: att.id, name: att.name,
            size: att.size, contentType: att.contentType
          });
        }
      }
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      mailData.body = result.status === Office.AsyncResultStatus.Succeeded
        ? result.value : "";
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
  if (mailData.attachments?.length > 0) {
    row.style.display = "flex";
    row.innerHTML = mailData.attachments.map((a) =>
      `<span class="att-badge">📎 ${escHtml(a.name)} (${fmtSize(a.size)})</span>`
    ).join("");
  }
}

// ─── GRAPHQL ──────────────────────────────────────────────────────────────────

async function heroQuery(query, variables) {
  const res = await fetch(HERO_GQL, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(variables ? { query, variables } : { query })
  });
  return res.json();
}

async function tryQuery(query, variables, dataKey) {
  try {
    const result = await heroQuery(query, variables);
    if (result.errors) {
      console.warn(`Query "${dataKey}":`, result.errors);
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

// ─── PROJECT SEARCH ───────────────────────────────────────────────────────────

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
    let projects = await tryQuery(`
      query ($search: String) {
        project_matches(search: $search, first: 50) {
          id project_nr
          customer { first_name last_name company_name }
          address { city }
          current_project_match_status { name }
        }
      }`, { search: term }, "project_matches");

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

    if (!projects?.length) {
      document.getElementById("projectResults").innerHTML =
        `<div class="state-msg">Keine Projekte gefunden</div>`;
      return;
    }

    renderProjects(projects);
  } catch (e) {
    console.error("searchProjects:", e);
    document.getElementById("projectResults").innerHTML =
      `<div class="state-msg">Fehler: ${e.message}</div>`;
  }
}

function matchesSearch(p, s) {
  const nr      = (p.project_nr || "").toLowerCase();
  const company = (p.customer?.company_name || "").toLowerCase();
  const name    = (`${p.customer?.first_name || ""} ${p.customer?.last_name || ""}`).toLowerCase();
  const city    = (p.address?.city || "").toLowerCase();
  return nr.includes(s) || company.includes(s) || name.includes(s) || city.includes(s);
}

function renderProjects(projects) {
  document.getElementById("projectResults").innerHTML = projects.map((p) => {
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
  document.querySelectorAll(".project-item").forEach((el) => el.classList.remove("selected"));
  event.currentTarget.classList.add("selected");
  const btn = document.getElementById("submitBtn");
  btn.disabled = false;
  btn.textContent = `📤 An ${nr} senden`;
}

// ─── PDF GENERATION ───────────────────────────────────────────────────────────

function generateEmailPdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const pageW = doc.internal.pageSize.getWidth();
  const margin = 15;
  const maxW = pageW - margin * 2;
  let y = 20;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.setTextColor(245, 155, 0);
  doc.text("E-Mail", margin, y);
  y += 8;

  doc.setDrawColor(226, 229, 234);
  doc.line(margin, y, pageW - margin, y);
  y += 6;

  doc.setFontSize(10);
  doc.setTextColor(95, 102, 114);

  for (const [label, value] of [
    ["Von:", `${mailData.fromName || ""} <${mailData.from || ""}>`],
    ["Betreff:", mailData.subject || ""],
    ["Datum:", mailData.date || ""],
  ]) {
    doc.setFont("helvetica", "bold");
    doc.text(label, margin, y);
    doc.setFont("helvetica", "normal");
    const lines = doc.splitTextToSize(value, maxW - 25);
    doc.text(lines, margin + 22, y);
    y += lines.length * 5 + 2;
  }

  if (mailData.attachments?.length > 0) {
    doc.setFont("helvetica", "bold");
    doc.text("Anhänge:", margin, y);
    doc.setFont("helvetica", "normal");
    const lines = doc.splitTextToSize(mailData.attachments.map((a) => a.name).join(", "), maxW - 25);
    doc.text(lines, margin + 22, y);
    y += lines.length * 5 + 2;
  }

  y += 4;
  doc.line(margin, y, pageW - margin, y);
  y += 7;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(26, 29, 35);

  for (const line of doc.splitTextToSize((mailData.body || "").substring(0, 8000), maxW)) {
    if (y > 275) { doc.addPage(); y = 20; }
    doc.text(line, margin, y);
    y += 4.5;
  }

  return doc.output("datauristring").split(",")[1];
}

// ─── SUBMIT ───────────────────────────────────────────────────────────────────

async function submitToHero() {
  if (!selectedProject || !apiKey) return;

  const btn = document.getElementById("submitBtn");
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner" style="width:14px;height:14px;border-width:2px;border-top-color:white"></span> Wird hochgeladen...`;

  try {
    showStatus("loading", "E-Mail als PDF wird generiert und hochgeladen...");
    const filename = `Email_${sanitizeFilename(mailData.subject)}_${new Date().toISOString().slice(0, 10)}.pdf`;
    const pdfBase64 = generateEmailPdf();
    await uploadFileToHero(filename, pdfBase64, "application/pdf");

    showStatus("success", `✅ PDF wurde Projekt ${selectedProject.nr} hinzugefügt!`);
  } catch (e) {
    showStatus("error", "❌ Fehler: " + e.message);
  } finally {
    btn.disabled = false;
    btn.textContent = `📤 An ${selectedProject.nr} senden`;
  }
}

// ─── UPLOAD ───────────────────────────────────────────────────────────────────

async function uploadFileToHero(filename, base64Content, contentType) {
  // Schritt 1: Datei → UUID
  const uploadRes = await fetch("/api/hero?upload=1", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ filename, content_base64: base64Content, content_type: contentType }),
  });

  const uploadData = await uploadRes.json();
  console.log("File upload response:", uploadData);

  const uuid = uploadData?.uuid || uploadData?.id || uploadData?.file_upload_uuid
    || uploadData?.data?.uuid || uploadData?.data?.id;

  if (!uuid) {
    throw new Error("Keine UUID erhalten: " + JSON.stringify(uploadData));
  }

  // Schritt 2: upload_document via GraphQL v8
  const gqlRes = await fetch("/api/hero?v8=1", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      query: `
        mutation ($uuid: String!, $targetId: Int!) {
          upload_document(
            target: project_match,
            target_id: $targetId,
            file_upload_uuid: $uuid,
            document: {}
          ) { id }
        }
      `,
      variables: { uuid, targetId: selectedProject.id },
    }),
  });

  const gqlData = await gqlRes.json();
  console.log("upload_document response:", gqlData);

  if (gqlData.errors) {
    throw new Error("upload_document: " + JSON.stringify(gqlData.errors));
  }

  return gqlData;
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function showStatus(type, message) {
  const bar = document.getElementById("statusBar");
  bar.className = `status-bar show ${type}`;
  bar.innerHTML = (type === "loading"
    ? `<div class="spinner" style="width:14px;height:14px;flex-shrink:0"></div>` : "")
    + message;
  if (type !== "loading") setTimeout(() => bar.classList.remove("show"), 6000);
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

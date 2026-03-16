/**
 * HERO Outlook Add-In
 * taskpane.js - Hauptlogik
 */

// HERO API Konfiguration
const HERO_CONFIG = {
    apiKey: 'ac_2RsFINoNFGI97t1jCvaiIZTl5DIKg1da',
    apiUrl: 'https://login.hero-software.de/api/v1',
    measure: 'PRJ'
};

// Globale Variablen
let currentItem = null;
let emailData = null;

/**
 * Office.js Initialize
 */
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('HERO Add-In geladen');
        loadEmailData();
    }
});

/**
 * E-Mail-Daten laden und anzeigen
 */
function loadEmailData() {
    currentItem = Office.context.mailbox.item;
    
    if (!currentItem) {
        showStatus('error', 'Keine E-Mail ausgewählt');
        return;
    }

    // E-Mail-Details anzeigen
    document.getElementById('emailFrom').textContent = currentItem.from?.displayName || currentItem.from?.emailAddress || '-';
    document.getElementById('emailSubject').textContent = currentItem.subject || '(Kein Betreff)';
    document.getElementById('emailDate').textContent = currentItem.dateTimeCreated?.toLocaleDateString('de-DE') || '-';

    // Anhänge anzeigen
    displayAttachments();

    // E-Mail-Body und weitere Daten abrufen
    getEmailBody();
    
    // Projekt-ID aus Betreff extrahieren (falls vorhanden)
    extractProjectIdFromSubject();
}

/**
 * E-Mail-Body abrufen
 */
function getEmailBody() {
    currentItem.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailData = {
                from: currentItem.from?.emailAddress || '',
                fromName: currentItem.from?.displayName || '',
                to: getRecipients(currentItem.to),
                cc: getRecipients(currentItem.cc),
                subject: currentItem.subject || '',
                body: result.value || '',
                dateReceived: currentItem.dateTimeCreated,
                hasAttachments: currentItem.attachments.length > 0,
                attachments: currentItem.attachments
            };
        }
    });
}

/**
 * Empfänger extrahieren
 */
function getRecipients(recipients) {
    if (!recipients || !recipients.length) return [];
    return recipients.map(r => ({
        name: r.displayName,
        email: r.emailAddress
    }));
}

/**
 * Anhänge anzeigen
 */
function displayAttachments() {
    const attachmentsList = document.getElementById('attachmentsList');
    const attachments = currentItem.attachments;

    if (!attachments || attachments.length === 0) {
        return;
    }

    let html = '<div style="margin-top: 10px;"><strong>Anhänge (' + attachments.length + '):</strong></div>';
    
    attachments.forEach(attachment => {
        const size = formatFileSize(attachment.size);
        html += `<div class="file-item">${attachment.name} (${size})</div>`;
    });

    attachmentsList.innerHTML = html;
}

/**
 * Projekt-ID aus Betreff extrahieren
 */
function extractProjectIdFromSubject() {
    const subject = currentItem.subject || '';
    
    // Suche nach Mustern wie [HERO-12345] oder [PRJ-2024-001]
    const patterns = [
        /\[HERO-(\d+)\]/i,
        /\[PRJ-(\d{4}-\d+)\]/i,
        /\[(\d+)\]/,
        /HERO[:\s-]*(\d+)/i,
        /PRJ[:\s-]*([\d-]+)/i
    ];

    for (const pattern of patterns) {
        const match = subject.match(pattern);
        if (match && match[1]) {
            document.getElementById('projectId').value = match[1];
            return;
        }
    }
}

/**
 * Zu HERO hochladen - Hauptfunktion
 */
async function uploadToHero() {
    const projectId = document.getElementById('projectId').value.trim();
    const documentType = document.getElementById('documentType').value;
    const notes = document.getElementById('notes').value.trim();

    // Validierung
    if (!projectId) {
        showStatus('error', 'Bitte Projekt-ID eingeben');
        return;
    }

    if (!emailData) {
        showStatus('error', 'E-Mail-Daten noch nicht geladen. Bitte warten...');
        return;
    }

    // UI blockieren
    setLoading(true);
    showStatus('loading', 'E-Mail wird hochgeladen...');

    try {
        // Schritt 1: E-Mail als Notiz/Kommentar erstellen
        const emailUploadResult = await createEmailNote(projectId, documentType, notes);

        if (!emailUploadResult.success) {
            throw new Error(emailUploadResult.error || 'Fehler beim Hochladen der E-Mail');
        }

        // Schritt 2: Anhänge hochladen (falls vorhanden)
        if (currentItem.attachments && currentItem.attachments.length > 0) {
            showStatus('loading', `Lade ${currentItem.attachments.length} Anhang/Anhänge hoch...`);
            await uploadAttachments(projectId);
        }

        // Erfolg!
        showStatus('success', '✅ Erfolgreich zu HERO hochgeladen!');
        
        // Optional: Formular zurücksetzen nach 3 Sekunden
        setTimeout(() => {
            document.getElementById('notes').value = '';
            document.getElementById('documentType').value = '';
        }, 3000);

    } catch (error) {
        console.error('Upload Fehler:', error);
        showStatus('error', '❌ Fehler: ' + error.message);
    } finally {
        setLoading(false);
    }
}

/**
 * E-Mail als Notiz in HERO erstellen
 */
async function createEmailNote(projectId, documentType, userNotes) {
    // E-Mail-Inhalt formatieren
    const emailContent = formatEmailForHero(documentType, userNotes);

    // API-Endpoint (muss an tatsächliche HERO API angepasst werden)
    const endpoint = `${HERO_CONFIG.apiUrl}/Projects/${projectId}/notes`;

    const payload = {
        title: `E-Mail: ${emailData.subject}`,
        content: emailContent,
        source: 'Outlook Add-In',
        type: documentType || 'email'
    };

    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${HERO_CONFIG.apiKey}`,
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        const result = await response.json();

        if (!response.ok) {
            return {
                success: false,
                error: result.message || `HTTP ${response.status}`
            };
        }

        return {
            success: true,
            data: result
        };

    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * E-Mail für HERO formatieren
 */
function formatEmailForHero(documentType, userNotes) {
    let content = '';

    // Header
    content += `Von: ${emailData.fromName} <${emailData.from}>\n`;
    content += `An: ${formatRecipientList(emailData.to)}\n`;
    
    if (emailData.cc && emailData.cc.length > 0) {
        content += `CC: ${formatRecipientList(emailData.cc)}\n`;
    }
    
    content += `Datum: ${emailData.dateReceived?.toLocaleString('de-DE') || ''}\n`;
    content += `Betreff: ${emailData.subject}\n`;
    
    if (documentType) {
        content += `Typ: ${documentType}\n`;
    }
    
    content += '\n';
    content += '─'.repeat(50) + '\n\n';

    // Body
    content += emailData.body;

    // User Notes
    if (userNotes) {
        content += '\n\n';
        content += '─'.repeat(50) + '\n';
        content += 'NOTIZEN:\n';
        content += userNotes;
    }

    // Anhänge
    if (emailData.hasAttachments) {
        content += '\n\n';
        content += '─'.repeat(50) + '\n';
        content += `ANHÄNGE (${emailData.attachments.length}):\n`;
        emailData.attachments.forEach((att, i) => {
            content += `${i + 1}. ${att.name} (${formatFileSize(att.size)})\n`;
        });
    }

    return content;
}

/**
 * Anhänge hochladen
 */
async function uploadAttachments(projectId) {
    const attachments = currentItem.attachments;
    const results = [];

    for (let i = 0; i < attachments.length; i++) {
        const attachment = attachments[i];
        
        try {
            showStatus('loading', `Lade Anhang ${i + 1}/${attachments.length} hoch: ${attachment.name}...`);
            
            // Anhang-Daten abrufen
            const attachmentData = await getAttachmentContent(attachment.id);
            
            // Zu HERO hochladen
            const uploadResult = await uploadAttachmentToHero(projectId, attachment, attachmentData);
            
            results.push({
                name: attachment.name,
                success: uploadResult.success
            });

        } catch (error) {
            console.error(`Fehler beim Hochladen von ${attachment.name}:`, error);
            results.push({
                name: attachment.name,
                success: false,
                error: error.message
            });
        }
    }

    return results;
}

/**
 * Anhang-Inhalt abrufen
 */
function getAttachmentContent(attachmentId) {
    return new Promise((resolve, reject) => {
        currentItem.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value.content); // Base64 encoded
            } else {
                reject(new Error('Fehler beim Abrufen des Anhangs'));
            }
        });
    });
}

/**
 * Anhang zu HERO hochladen
 */
async function uploadAttachmentToHero(projectId, attachment, base64Data) {
    // API-Endpoint für Dokument-Upload (muss angepasst werden)
    const endpoint = `${HERO_CONFIG.apiUrl}/Projects/${projectId}/documents`;

    // Konvertiere Base64 zu Blob
    const blob = base64ToBlob(base64Data, attachment.contentType);

    // FormData erstellen
    const formData = new FormData();
    formData.append('file', blob, attachment.name);
    formData.append('description', `Anhang aus E-Mail: ${emailData.subject}`);
    formData.append('source', 'Outlook Add-In');

    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${HERO_CONFIG.apiKey}`,
            },
            body: formData
        });

        const result = await response.json();

        return {
            success: response.ok,
            data: result
        };

    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Hilfsfunktionen
 */
function formatRecipientList(recipients) {
    if (!recipients || recipients.length === 0) return '';
    return recipients.map(r => `${r.name} <${r.email}>`).join(', ');
}

function formatFileSize(bytes) {
    if (!bytes) return '0 B';
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
}

function base64ToBlob(base64, contentType = '') {
    const byteCharacters = atob(base64);
    const byteArrays = [];

    for (let offset = 0; offset < byteCharacters.length; offset += 512) {
        const slice = byteCharacters.slice(offset, offset + 512);
        const byteNumbers = new Array(slice.length);
        
        for (let i = 0; i < slice.length; i++) {
            byteNumbers[i] = slice.charCodeAt(i);
        }
        
        const byteArray = new Uint8Array(byteNumbers);
        byteArrays.push(byteArray);
    }

    return new Blob(byteArrays, { type: contentType });
}

/**
 * UI Helper
 */
function showStatus(type, message) {
    const statusDiv = document.getElementById('statusMsg');
    statusDiv.className = `status ${type}`;
    
    if (type === 'loading') {
        statusDiv.innerHTML = `<span class="spinner"></span>${message}`;
    } else {
        statusDiv.innerHTML = message;
    }
    
    statusDiv.style.display = 'block';

    // Auto-hide nach 5 Sekunden (außer bei loading)
    if (type !== 'loading') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }
}

function setLoading(isLoading) {
    const btn = document.getElementById('uploadBtn');
    btn.disabled = isLoading;
    btn.textContent = isLoading ? 'Wird hochgeladen...' : 'Zu HERO hochladen';
}

function showSettings() {
    alert('Einstellungen (TODO): API-Key, Standard-Projekt, etc.');
    // TODO: Settings-Dialog implementieren
}

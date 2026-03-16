# HERO Mail Sync – Outlook Add-In

## Übersicht

Outlook Add-In für Kluge Elektriker. Ermöglicht das direkte Zuordnen von E-Mails (inkl. Anhänge) zu HERO-Projekten – direkt aus Outlook heraus.

## Features

- 📧 Zeigt die aktuelle E-Mail (Betreff, Absender, Anhänge) im Seitenpanel
- 🔍 Live-Suche nach HERO-Projekten (Projektnr., Kundenname)
- 📝 Schreibt Mail-Inhalt als Logbuch-Eintrag ins gewählte Projekt
- 📎 Lädt Anhänge zum Projekt hoch (via GraphQL API)
- ⚙️ API-Key wird lokal gespeichert

## Dateien

```
hero-outlook-addin/
├── manifest.xml      ← Add-In Manifest (für Outlook-Deployment)
├── taskpane.html     ← Hauptdatei (HTML + CSS + JS in einer Datei)
└── README.md         ← Diese Anleitung
```

## Voraussetzungen

1. **HERO GraphQL API-Key** – kostenlos beim HERO Support anfragen
2. **Microsoft 365 Konto** (Outlook im Web oder Desktop)
3. **Webhosting** für die `taskpane.html` (HTTPS erforderlich!)

## Deployment

### Schritt 1: Hosting

Die `taskpane.html` muss unter HTTPS erreichbar sein. Optionen:

- **Eigener Webserver** (z.B. Strato VPS, DigitalOcean)
- **GitHub Pages** (kostenlos, HTTPS automatisch)
- **Azure Static Web Apps** (ideal für Microsoft-Umgebung)

Beispiel: `https://tools.kluge-elektriker.de/hero-addin/taskpane.html`

### Schritt 2: manifest.xml anpassen

In `manifest.xml` alle Vorkommen von `https://DEINE-DOMAIN.de/hero-addin/` ersetzen durch die tatsächliche URL wo die Dateien gehostet sind.

### Schritt 3: Add-In in Outlook installieren

**Option A: Sideloading (zum Testen)**
1. Outlook im Web öffnen (outlook.office.com)
2. Zahnrad → "Alle Outlook-Einstellungen anzeigen"
3. "E-Mail" → "Aktionen anpassen" → "Add-Ins"
4. "Meine Add-Ins" → "Benutzerdefinierte Add-Ins hinzufügen" → "Aus Datei hinzufügen"
5. Die `manifest.xml` hochladen

**Option B: Admin-Deployment (für alle Mitarbeiter)**
1. Microsoft 365 Admin Center öffnen
2. Einstellungen → "Integrierte Apps"
3. "Benutzerdefinierte Apps hochladen" → manifest.xml
4. Allen Nutzern oder bestimmten Gruppen zuweisen

### Schritt 4: Erster Start

1. In Outlook eine E-Mail öffnen
2. Den "HERO" Button in der Toolbar klicken
3. API-Key in den Einstellungen eingeben
4. Verbindung testen – fertig!

## Nutzung

1. E-Mail in Outlook öffnen
2. "An HERO senden" Button klicken → Seitenpanel öffnet sich
3. Projekt suchen (Projektnr. oder Kundenname)
4. Projekt anklicken
5. "An [Projekt] senden" klicken
6. Mail-Inhalt wird ins Logbuch geschrieben, Anhänge hochgeladen

## Wichtige Hinweise

### GraphQL Mutations verifizieren

Die folgenden Mutations müssen mit dem tatsächlichen HERO Schema abgeglichen werden:

1. **Logbuch-Eintrag:** `add_logbook_entry` – Laut HERO Doku vorhanden
2. **Datei-Upload:** `upload_project_file` – Muss im Schema verifiziert werden

→ Mit **Insomnia** das Schema explorieren:
- URL: `https://login.hero-software.de/api/external/v7/graphql`
- Header: `Authorization: Bearer DEIN_API_KEY`
- Introspection Query ausführen und nach Upload/File Mutations suchen

### Projekt-Suche

Die Suche nutzt den `search` Parameter der `project_matches` Query. Falls die HERO API-Version diesen Parameter nicht unterstützt, greift ein Fallback der alle Projekte lädt und clientseitig filtert (funktioniert gut bei < 500 Projekten).

### CORS

Die HERO API muss Cross-Origin Requests zulassen. Falls CORS-Fehler auftreten, muss ein Proxy-Server dazwischengeschaltet werden (z.B. eine kleine Node.js/Express App auf eurem Server).

## Anpassungen

### Gewerk ändern
Im Code wird kein Gewerk festgelegt, da das Add-In nur bestehenden Projekten zuordnet (keine neuen Projekte anlegt).

### Body-Zeichenlimit
Standard: 3000 Zeichen. Anpassbar in `buildLogMessage()` in der `taskpane.html`.

### Styling
Farben und Design können in den CSS-Variablen oben in der `taskpane.html` angepasst werden.

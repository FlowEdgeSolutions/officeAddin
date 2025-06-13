# 📧 MailAider AI – Outlook Add-In

**MailAider AI** ist ein intelligentes Outlook-Add-in, das mithilfe von GPT-basierten KI-Modellen eingehende E-Mails analysiert, zusammenfasst, beantwortet oder übersetzt. Ziel ist es, den E-Mail-Alltag durch Automatisierung und smarte Textvorschläge deutlich effizienter zu gestalten – sowohl für Einzelpersonen als auch für Unternehmen.

---

## 🚀 Features

- ✉️ **E-Mail-Zusammenfassung:** Kompakte Stichpunkte auf Basis des Mail-Inhalts
- 🤖 **KI-gestützte Antwortvorschläge:** Ton, Länge und Anrede individuell konfigurierbar
- 🌐 **Übersetzungen:** Inhalte in andere Sprachen übertragen
- 🧠 **Freier Modus:** Eigene Prompts an Azure OpenAI oder lokale Ollama-Instanz senden
- 🌗 **Dark Mode Toggle** für bessere Lesbarkeit
- 🔒 **Zwei Verarbeitungsmodi:**
  - Azure OpenAI (schnell, Cloud)
  - Ollama (datenschutzfreundlich, lokal) - Einrichtung ist noch nicht erfolgt!

---

## 🛠️ Technologien

- Office Add-in (MailApp)
- HTML/CSS/JS (kein Framework)
- Office.js (Microsoft 365)
- Azure OpenAI GPT-4
- Optional: Ollama + Nous Hermes 2 (lokale GPT-Alternative)

---

## 📦 Setup & Deployment

### Voraussetzungen

- Microsoft 365 Konto mit Outlook
- Azure Static Web App (z. B. `https://witty-stone-066af150f.6.azurestaticapps.net`)
- Zugriff auf Azure OpenAI API oder lokale Ollama-Instanz

### Deployment-Schritte

1. **Build & Upload**
   - Statische HTML-Dateien (inkl. `taskpane.html`, `function.html`) in Azure Static Web App deployen.

2. **Manifest-XML aktualisieren**
   - `SourceLocation` und `resid`-URLs an die eigene Domain anpassen
   - Beispiel: `https://orange-dune-0a75eea0f.6.azurestaticapps.net/taskpane.html`

3. **Sideload in Outlook**
   - `manifest.xml` in Outlook als benutzerdefiniertes Add-in laden (Outlook Desktop/Web)

---

## 🧪 Testen in Outlook

### Outlook Web (Office Online)

1. Navigiere zu [https://outlook.office.com](https://outlook.office.com)
2. Klicke oben rechts auf ⚙️ → **Alle Outlook-Einstellungen anzeigen**
3. Gehe zu **E-Mail > Add-Ins > Eigene Add-Ins**
4. Wähle **Ein benutzerdefiniertes Add-In hinzufügen > Aus Datei**
5. Wähle deine `manifest.xml`

### Outlook Desktop (Windows)

1. Öffne Outlook
2. Aktiviere Entwickler-Tools (Datei → Optionen → Add-ins)
3. Lade `manifest.xml` über die Entwicklertools

---

## ⚠️ Sicherheitshinweise

> **Achtung:** Die Datei enthält aktuell einen **fest eingebetteten Azure OpenAI API-Key**.  
Das ist **nicht sicher** für produktive Umgebungen! Bitte:

- Entferne API-Keys aus dem Frontend.
- Verwende einen Backend-Proxy (Azure Function, Node.js etc.) zur sicheren Kommunikation.
- Für produktive Deployments: nutze Authentifizierung (z. B. MSAL, Entra ID) und Token-Validierung.

---

## 📄 Lizenz

Dieses Projekt ist derzeit nicht öffentlich lizenziert. Für kommerzielle Nutzung, Weitergabe oder Veröffentlichung ist die Zustimmung des Autors erforderlich.

© 2025 [FlowEdge Solutions](https://flowedge.de)

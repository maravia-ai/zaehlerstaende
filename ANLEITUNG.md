# Zählerstände-App – Einrichtung

Web-App für iPhone und Browser zur Erfassung der Zählerstände (Gas, Wasser, Strom) für
**Burgacker 39** und **Leipziger Str. 7**. Datenquelle ist die Excel-Datei
`Zählerstände.xlsx` auf deinem OneDrive (Microsoft 365).

## Dateien

| Datei | Zweck |
|---|---|
| `index.html` | Oberfläche |
| `app.js` | Logik (Excel lesen/schreiben, Verbrauch, Fotos, Kalender) |
| `config.js` | Einstellungen (Client-ID, Tenant-ID, Dateipfad, Fotoordner, Erinnerung) |
| `msal-browser.min.js` | Microsoft-Anmeldebibliothek (v3.28.1, lokal mitgeliefert) |
| `manifest.webmanifest`, `icon-180.png`, `icon-512.png` | App-Symbol für den Home-Bildschirm |
| `vorlage/Zählerstände.xlsx` | Excel-Vorlage (einmalig ins OneDrive kopieren, enthält keine persönlichen Daten) |

## 1. Excel-Datei ablegen

`vorlage/Zählerstände.xlsx` in dein OneDrive nach `1 - Dirk/06 Immobilien/Zählerstände.xlsx` kopieren.
Ein anderer Ort ist möglich, dann `filePath` und `fotoOrdner` in `config.js` anpassen.

Blätter:
- **Ablesungen**: Excel-Tabelle „Ablesungen“, wird von der App befüllt.
- **Zähler**: Zählernummern in den gelben Zellen eintragen (oder später beim ersten Erfassen in der App).
- **Übersicht**: Verbrauch je Monat und Jahr, berechnet automatisch.
- **Hinweise**: kurze Erklärung.

## 2. GitHub Pages einschalten (einmalig)

Die App wird wie die Medikamente-App über GitHub Pages bereitgestellt.

1. https://github.com/maravia-ai/zaehlerstaende → **Settings → Pages**
2. **Source:** „Deploy from a branch“, **Branch:** `main`, Ordner `/ (root)` → **Save**
3. Nach 1–2 Minuten ist die App erreichbar unter
   `https://maravia-ai.github.io/zaehlerstaende/`

Änderungen werden automatisch veröffentlicht, sobald sie auf `main` liegen.

## 3. App-Registrierung in Microsoft Entra ID

1. https://entra.microsoft.com → **Identität → Anwendungen → App-Registrierungen → Neue Registrierung**
2. Name: `Zählerstände-App`
3. Unterstützte Kontotypen: **Nur Konten in diesem Organisationsverzeichnis (nur maravia)**
4. Umleitungs-URI: Plattform **Single-Page-Anwendung (SPA)**,
   `https://maravia-ai.github.io/zaehlerstaende/` → **Registrieren**
5. Unter **Authentifizierung** zusätzlich `https://maravia-ai.github.io/zaehlerstaende/index.html` eintragen.
6. **API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen**:
   `Files.ReadWrite`, `Calendars.ReadWrite` (`User.Read` ist schon vorhanden).
   Optional: **Administratorzustimmung für maravia erteilen**.
7. **Anwendungs-ID (Client)** von der Übersichtsseite in `config.js` bei `clientId` eintragen.
   Die Mandanten-ID ist bereits eingetragen (dieselbe wie bei der Medikamente-App).

Die geänderte `config.js` muss auf `main` liegen (Claude pusht sie, wenn du die ID schickst).

Alternative: In der bestehenden Registrierung „Medikamente-App“ nur die beiden neuen
Umleitungs-URIs ergänzen und deren Client-ID verwenden. Die Berechtigungen sind dort identisch.

## 4. Auf dem iPhone installieren

1. Adresse in **Safari** öffnen
2. **Teilen → Zum Home-Bildschirm**
3. App starten → **Mit Microsoft anmelden**
4. Einmal **Monatliche Erinnerung einrichten** tippen (Serientermin am 1. jedes Monats, 09:00 Uhr,
   im Microsoft-365-Kalender; Tag und Uhrzeit in `config.js` änderbar). Erneutes Tippen ersetzt den Termin.

## Bedienung

- **Übersicht:** pro Gebäude die drei Zähler mit letztem Stand und letztem Verbrauch.
  „offen“ = im laufenden Monat noch nicht abgelesen, „erfasst“ = erledigt.
- **Zähler antippen → erfassen:** Datum (Standard heute, rückdatieren möglich), Zählerstand
  (Komma oder Punkt als Dezimaltrennzeichen), optional Foto und Notiz.
  Die App zeigt vor dem Speichern den Verbrauch seit der letzten Ablesung.
- **Zweirichtungszähler Burgacker 39 (PV):** Beim Strom gibt es zwei Felder, **1.8.0 Bezug** (aus dem Netz)
  und **2.8.0 Einspeisung** (ins Netz); beide sind Pflicht. Zu jedem Wert kann ein eigenes Foto als Nachweis
  aufgenommen werden (Dateiname endet auf `_1-8-0.jpg` bzw. `_2-8-0.jpg`), verlinkt in der jeweiligen Excel-Zeile. In Excel entstehen zwei Zeilen: Medium „Strom“ (1.8.0)
  und „Strom Einspeisung“ (2.8.0), mit derselben Zählernummer. Hinweis: 1.8.0 ist nur der Netzbezug,
  nicht der Gesamtverbrauch des Hauses (der selbst verbrauchte PV-Strom wird vom Zähler nicht erfasst).
- **Plausibilitätsprüfung:** Ist der Stand kleiner als zuvor oder der Tagesverbrauch mehr als dreimal so hoch
  wie im Schnitt der letzten 6 Ablesungen, erscheint eine Warnung; der Button heißt dann „Trotzdem speichern“.
- **Zwischenablesungen** sind jederzeit möglich (z. B. bei Mieterwechsel), einfach erneut erfassen.
- **Zählerwechsel:** erst den Endstand des alten Zählers erfassen, dann beim neuen Zähler die neue
  Zählernummer eingeben. Für diese Ablesung wird kein Verbrauch berechnet; die Nummer im Blatt „Zähler“
  wird aktualisiert.
- **Fotos** werden verkleinert (max. 1600 px) im OneDrive-Ordner `1 - Dirk/06 Immobilien/Zählerstände Fotos` gespeichert und in
  Spalte „Foto“ verlinkt.
- **Korrekturen** direkt in Excel. Die App berechnet bereits gespeicherte Verbrauchswerte nicht neu.

## Demo

`https://maravia-ai.github.io/zaehlerstaende/?demo` zeigt die App mit Beispieldaten, ohne Anmeldung und ohne zu speichern.

## Einschränkungen

- Spaltenreihenfolge und Überschriften der Tabelle „Ablesungen“ sowie die Kopfzeile im Blatt „Zähler“
  müssen erhalten bleiben. Die App prüft sie und meldet Abweichungen.
- Microsoft-Anmeldungen im Browser bleiben in der Regel ca. 24 Stunden ohne Rückfrage gültig.

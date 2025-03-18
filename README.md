# Windows Event Analyzer mit Claude 3.7 Sonnet

Dieses PowerShell-Tool analysiert Windows-Ereignisprotokolle mit Hilfe von Claude 3.7 Sonnet (via OpenRouter) und stellt die Ergebnisse in einer benutzerfreundlichen GUI dar.

## Funktionen

- Erfasst die letzten 50 Ereignisse aus dem Windows-Systemprotokoll
- Sendet die Daten an Claude 3.7 Sonnet via OpenRouter fuer eine KI-gestuetzte Analyse
- Zeigt die Analyse-Ergebnisse in einer uebersichtlichen GUI mit Markdown-Formatierung an
- Speichert Analysen als Textdateien oder Markdown-Dokumente
- Bietet eine nutzerfreundliche Oberflaeche mit Menueleiste und Statusanzeige

## Voraussetzungen

- Windows mit PowerShell 5.1 oder hoeher
- Internetverbindung
- OpenRouter API-Schluessel (mit Zugriff auf Claude 3.7 Sonnet)

## Installation

1. Lade die Dateien herunter oder klone das Repository
2. API-Schluessel konfigurieren (zwei Moeglichkeiten):
   - **Option A**: Bearbeite die `.env` Datei im Projektverzeichnis und trage deinen OpenRouter API-Schluessel ein
   - **Option B**: Fuehre das Skript aus und gib deinen API-Schluessel ein, wenn du danach gefragt wirst
3. Der API-Schluessel wird je nach Methode entweder in der lokalen `.env` Datei oder in `%USERPROFILE%\Documents\EventAnalyzer\.env` gespeichert

Das Tool erstellt automatisch einen Ordner `EventAnalyzer` im Dokumente-Verzeichnis des Benutzers zum Speichern von Analysen.

## Verwendung

### Hauptversion mit Claude 3.7 Sonnet (erfordert OpenRouter API-Schluessel)

1. Oeffne PowerShell mit administrativen Rechten (Rechtsklick auf PowerShell > "Als Administrator ausfuehren")
2. Navigiere zum Verzeichnis mit dem Skript:
   ```
   cd Pfad\zum\Verzeichnis
   ```
3. Fuehre das Skript aus:
   ```
   .\EventAnalyzer.ps1
   ```
4. Das Skript:
   - Sucht zuerst nach einer lokalen `.env` Datei im Projektverzeichnis
   - Falls keine lokale Datei gefunden wird, verwendet es die Datei in `%USERPROFILE%\Documents\EventAnalyzer\.env`
   - Fragt nur nach einem API-Schluessel, wenn in keiner der beiden Dateien ein Schluessel gefunden wird

5. Die Anwendung:
   - Liest die letzten 50 Systemereignisse
   - Sendet sie an Claude 3.7 Sonnet zur Analyse
   - Zeigt das Ergebnis in einem Fenster mit Markdown-Formatierung an
   - Erlaubt das Speichern der Analyse ueber das Menue "Datei" > "Analyse speichern..."

### Demo-Version (kein API-Schluessel erforderlich)

Wenn Sie das Tool ohne API-Schluessel testen moechten:

1. Oeffne PowerShell mit administrativen Rechten
2. Navigiere zum Verzeichnis mit dem Skript
3. Fuehre die Demo-Version aus:
   ```
   .\DemoEventAnalyzer.ps1
   ```
4. Die Demo-Version:
   - Sammelt 10 tatsaechliche Systemereignisse
   - Zeigt eine vordefinierte Analyse an (keine echte KI-Analyse)
   - Enthaelt eine Statusanzeige, die den Demo-Modus kennzeichnet

## Konfiguration anpassen

Im Skript koennen verschiedene Parameter angepasst werden:

- **Ereignisprotokoll-Einstellungen:**
  - `$maxEvents`: Anzahl der zu analysierenden Ereignisse (Standard: 50)
  - `$logName`: Zu analysierendes Ereignisprotokoll (Standard: "System", Alternativen: "Application", "Security")

- **API-Einstellungen:**
  - `$apiUrl`: Die OpenRouter API-URL
  - `$aiModell`: Das zu verwendende KI-Modell
    - Standard: "anthropic/claude-3.7-sonnet"
    - Alternative: "anthropic/claude-3.7-sonnet:thinking" (mit Denkprozess)

## Fehlerbehebung

- **Zugriffsfehler**: Stelle sicher, dass PowerShell mit administrativen Rechten ausgefuehrt wird
- **API-Fehler**: Ueberpruefe, ob dein API-Schluessel korrekt ist und du ueber ausreichend Guthaben verfuegst
- **Verbindungsprobleme**: Ueberpruefe deine Internetverbindung und die korrekte API-URL
- **Encoding-Probleme**: Bei Zeichenproblemen in der Ausgabe, stelle sicher dass die Konsole UTF-8 unterstuetzt
- **Speicherfehler**: Pruefe, ob der Benutzer Schreibrechte im Verzeichnis `%USERPROFILE%\Documents\EventAnalyzer` hat

## Lizenz

Dieses Projekt steht unter der MIT-Lizenz - siehe die [LICENSE](LICENSE) Datei für Details.

Hier ist ein kompletter PowerShell-Prompt, der alle bisherigen Schritte zusammenfasst – von der Ereignisdaten-Erfassung über die Analyse via OpenRouter bis hin zur Anzeige der Ergebnisse in einer GUI. Dabei stellen wir sicher, dass das Ausgabeformat korrekt dargestellt wird, indem wir z. B. ein mehrzeiliges Textfeld (Multiline TextBox) mit Scrollbars verwenden.

---

```powershell
# --- Konfiguration ---
$apiKey = "DEIN_API_SCHLÜSSEL"  # Deinen OpenRouter API-Schlüssel hier einfügen
$apiUrl = "https://api.openrouter.ai/v1/chat/completions"  # Überprüfe die korrekte URL in der OpenRouter-Dokumentation
$maxEvents = 50  # Anzahl der zu lesenden Ereignisdaten

# --- Ereignisdaten erfassen ---
try {
    $logData = Get-WinEvent -LogName System -MaxEvents $maxEvents |
        Select-Object Id, LevelDisplayName, TimeCreated, Message |
        ConvertTo-Json -Depth 3
} catch {
    Write-Error "Fehler beim Auslesen der Ereignisdaten: $_"
    exit
}

# --- Anfrage an OpenRouter API vorbereiten ---
$body = @{
    "model" = "gpt-4"  # oder "gpt-3.5-turbo", je nach Bedarf
    "messages" = @(
        @{
            "role" = "system"
            "content" = "Analysiere die folgenden Windows-Ereignisdaten und gib eine verständliche Zusammenfassung sowie mögliche Ursachen von Fehlern an."
        },
        @{
            "role" = "user"
            "content" = $logData
        }
    )
} | ConvertTo-Json

$headers = @{
    "Authorization" = "Bearer $apiKey"
    "Content-Type"  = "application/json"
}

# --- Anfrage senden und Antwort empfangen ---
try {
    $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $body
    # Hier nehmen wir an, dass die Antwort in $response.choices[0].message.content enthalten ist.
    $analyseErgebnis = $response.choices[0].message.content
} catch {
    $analyseErgebnis = "Fehler beim Abrufen der Analyse: $_"
}

# --- GUI erstellen zur Darstellung der Analyse ---
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Ereignis-Analyse"
$form.Size = New-Object System.Drawing.Size(800,600)

# TextBox erstellen für die Anzeige der Analyse (Mehrzeilig, mit Scrollbar)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$textBox.Dock = "Fill"
$textBox.Font = New-Object System.Drawing.Font("Consolas",10)
$textBox.WordWrap = $true

# Sicherstellen, dass Zeilenumbrüche korrekt angezeigt werden
$textBox.Text = $analyseErgebnis

$form.Controls.Add($textBox)
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)
```

---

### Erklärung

- **Ereignisdaten-Erfassung:**  
  Mit `Get-WinEvent` werden die letzten 50 Einträge aus dem Systemlog gelesen und in JSON konvertiert. Dabei sorgt `-Depth 3` dafür, dass auch verschachtelte Objekte vollständig serialisiert werden.

- **API-Integration:**  
  Das Skript baut eine JSON-Anfrage, die den Systemkontext (System Message) und die Ereignisdaten (User Message) enthält. Mit `Invoke-RestMethod` wird diese Anfrage an OpenRouter gesendet. Die Antwort wird aus der erwarteten Struktur (`$response.choices[0].message.content`) extrahiert.

- **GUI-Ausgabe:**  
  Ein Windows Forms-Fenster mit einer mehrzeiligen TextBox wird erstellt, um die Analyseergebnisse anzuzeigen. Die TextBox ist so konfiguriert, dass sie Zeilenumbrüche und Scrollbars unterstützt – so wird das Format korrekt wiedergegeben.

Dieses Skript benötigt nur PowerShell und die in Windows integrierten .NET-Bibliotheken, sodass keine zusätzlichen Installationen erforderlich sind. Passe den API-Schlüssel, die API-URL und ggf. weitere Parameter nach deinen Bedürfnissen an.

Falls du noch Fragen hast oder weitere Anpassungen benötigst, lass es mich wissen!
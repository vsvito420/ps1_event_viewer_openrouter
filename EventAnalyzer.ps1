# --- Windows Event Analyzer mit OpenRouter AI ---
# Analysiert Windows-Ereignisprotokolle mit Hilfe von Claude 3.7 Sonnet via OpenRouter

# --- Konfiguration ---
# Setze die Ausgabecodierung auf UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Pfade fuer Konfiguration und Speicherung
$appDir = "$env:USERPROFILE\Documents\EventAnalyzer"
$localEnvFile = ".\.env"  # Lokale .env Datei im aktuellen Verzeichnis
$appDirEnvFile = "$appDir\.env"  # .env Datei im Applikationsverzeichnis
$outputDir = $appDir

# Erstelle Verzeichnis falls nicht vorhanden
if (-not (Test-Path -Path $appDir)) {
    New-Item -ItemType Directory -Path $appDir | Out-Null
    Write-Host "Verzeichnis $appDir wurde erstellt." -ForegroundColor Green
}

# Bestimme welche .env Datei verwendet werden soll
if (Test-Path -Path $localEnvFile) {
    $configFile = $localEnvFile
    Write-Host "Lokale .env Datei im aktuellen Verzeichnis gefunden." -ForegroundColor Cyan
}
else {
    $configFile = $appDirEnvFile
    Write-Host "Verwende .env Datei im Applikationsverzeichnis: $appDirEnvFile" -ForegroundColor Cyan
}

# --- Funktion zum Lesen und Schreiben der .env Datei ---
function Get-EnvVariable {
    param (
        [string]$Key,
        [string]$FilePath
    )
    
    if (Test-Path $FilePath) {
        $envContent = Get-Content $FilePath
        foreach ($line in $envContent) {
            if ($line -match "^$Key=(.*)$") {
                return $matches[1]
            }
        }
    }
    return $null
}

function Set-EnvVariable {
    param (
        [string]$Key,
        [string]$Value,
        [string]$FilePath
    )
    
    $newLine = "$Key=$Value"
    
    if (Test-Path $FilePath) {
        $envContent = Get-Content $FilePath
        $updated = $false
        
        $newContent = @()
        foreach ($line in $envContent) {
            if ($line -match "^$Key=") {
                $newContent += $newLine
                $updated = $true
            }
            else {
                $newContent += $line
            }
        }
        
        if (-not $updated) {
            $newContent += $newLine
        }
        
        $newContent | Set-Content $FilePath
    }
    else {
        $newLine | Set-Content $FilePath
    }
}

# API Konfiguration
$apiUrl = "https://openrouter.ai/api/v1/chat/completions"  # OpenRouter API-URL
# Verfügbare Modelle
$availableModels = @{
    "Claude 3.7 Sonnet"            = "anthropic/claude-3.7-sonnet"
    "Claude 3.7 Sonnet (thinking)" = "anthropic/claude-3.7-sonnet:thinking"
    "GPT-4o"                       = "openai/gpt-4o"
    "GPT-4"                        = "openai/gpt-4"
    "Gemini 1.5 Pro"               = "google/gemini-1.5-pro"
}

# Standard-Modell ist Claude 3.7 Sonnet
$aiModell = $availableModels["Claude 3.7 Sonnet"]

# Modell-Auswahl anzeigen
Write-Host "`nVerfügbare KI-Modelle:" -ForegroundColor Cyan
$i = 1
$modelList = @{}
foreach ($model in $availableModels.Keys) {
    Write-Host "[$i] $model ($($availableModels[$model]))" -ForegroundColor Yellow
    $modelList[$i] = $model
    $i++
}

$modelChoice = Read-Host "`nWähle ein Modell (1-$($availableModels.Count)) [Standard: 1]"
if ($modelChoice -ne "" -and $modelChoice -match '^\d+$' -and [int]$modelChoice -ge 1 -and [int]$modelChoice -le $availableModels.Count) {
    $selectedModel = $modelList[[int]$modelChoice]
    $aiModell = $availableModels[$selectedModel]
    Write-Host "Ausgewähltes Modell: $selectedModel ($aiModell)" -ForegroundColor Green
}
else {
    Write-Host "Verwende Standard-Modell: Claude 3.7 Sonnet ($aiModell)" -ForegroundColor Green
}

# Test der Internetverbindung
Write-Host "Teste die Internetverbindung..." -ForegroundColor Yellow
try {
    $testConnection = Test-Connection -ComputerName "google.com" -Count 1 -Quiet
    if ($testConnection) {
        Write-Host "Internetverbindung OK" -ForegroundColor Green
    }
    else {
        Write-Host "Keine Internetverbindung verfügbar. Die API kann nicht erreicht werden." -ForegroundColor Red
        Write-Host "Starte Demo-Modus stattdessen..." -ForegroundColor Yellow
        & "$PSScriptRoot\DemoEventAnalyzer.ps1"
        exit
    }
}
catch {
    Write-Host "Fehler beim Prüfen der Internetverbindung: $_" -ForegroundColor Red
    Write-Host "Starte Demo-Modus stattdessen..." -ForegroundColor Yellow
    & "$PSScriptRoot\DemoEventAnalyzer.ps1"
    exit
}

# API-Schluessel aus .env Datei laden oder abfragen
$apiKey = Get-EnvVariable -Key "OPENROUTER_API_KEY" -FilePath $configFile

if ($null -eq $apiKey -or $apiKey -eq "") {
    Write-Host "`n--- OpenRouter API-Schluessel Konfiguration ---" -ForegroundColor Cyan
    Write-Host "Ein OpenRouter API-Schluessel ist erforderlich fuer die Analyse mit Claude 3.7 Sonnet."
    Write-Host "Der Schluessel wird in $configFile gespeichert und verwendet." -ForegroundColor Yellow
    
    $apiKey = Read-Host "Bitte gib deinen OpenRouter API-Schluessel ein"
    
    if ($apiKey -ne "") {
        Set-EnvVariable -Key "OPENROUTER_API_KEY" -Value $apiKey -FilePath $configFile
        Write-Host "API-Schluessel erfolgreich gespeichert." -ForegroundColor Green
    }
    else {
        Write-Error "Kein API-Schluessel eingegeben. Das Skript wird beendet."
        exit
    }
}

# Ereignis-Konfiguration
$maxEvents = 50  # Anzahl der zu lesenden Ereignisdaten
$logName = "System"  # Ereignisprotokoll (System, Application, Security)

# --- Funktion zum Erfassen der Ereignisdaten ---
function Get-EventLogData {
    param (
        [string]$LogName = "System",
        [int]$MaxEvents = 50
    )
    
    Write-Host "Erfasse Ereignisdaten aus $LogName-Protokoll..."
    
    try {
        $events = Get-WinEvent -LogName $LogName -MaxEvents $MaxEvents -ErrorAction Stop | 
        Select-Object Id, LevelDisplayName, TimeCreated, Message
        
        Write-Host "$($events.Count) Ereignisse erfolgreich gesammelt."
        return $events | ConvertTo-Json -Depth 3
    }
    catch {
        Write-Error "Fehler beim Auslesen der Ereignisdaten: $_"
        return $null
    }
}

# --- Ereignisdaten erfassen ---
$logData = Get-EventLogData -LogName $logName -MaxEvents $maxEvents
if ($null -eq $logData) {
    exit
}

# --- Funktion zur API-Anfrage ---
function Get-AIAnalysis {
    param (
        [string]$LogData,
        [string]$ApiUrl,
        [string]$ApiKey,
        [string]$Model
    )
    
    Write-Host "Sende Daten an $Model zur Analyse..."
    
    $systemPrompt = @"
Analysiere die folgenden Windows-Ereignisdaten und gib eine verstaendliche Zusammenfassung mit folgenden Abschnitten:
1. Uebersicht: Anzahl und Arten der Ereignisse
2. Wichtige Ereignisse: Hervorheben kritischer oder ungewoehnlicher Eintraege
3. Fehleranalyse: Moegliche Ursachen fuer Fehler oder Warnungen
4. Empfehlungen: Konkrete Handlungsempfehlungen basierend auf den Ereignissen
5. Zusammenfassung: Allgemeiner Systemzustand und wichtigste Punkte

Formatiere die Ausgabe mit Markdown fuer bessere Lesbarkeit.
"@
    
    $body = @{
        "model"    = $Model
        "messages" = @(
            @{
                "role"    = "system"
                "content" = $systemPrompt
            },
            @{
                "role"    = "user"
                "content" = $LogData
            }
        )
    } | ConvertTo-Json -Depth 4
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type"  = "application/json"
    }
    
    try {
        # API-Anfrage mit curl-ähnlicher Protokollierung
        Write-Host "`nSende API-Anfrage an OpenRouter:" -ForegroundColor Yellow
        Write-Host "URL: $ApiUrl" -ForegroundColor DarkGray
        Write-Host "Modell: $Model" -ForegroundColor DarkGray
        Write-Host "API-Schlüssel: $(if ($ApiKey.Length -gt 8) { $ApiKey.Substring(0, 4) + "..." + $ApiKey.Substring($ApiKey.Length - 4) } else { "Ungültig oder zu kurz" })" -ForegroundColor DarkGray
        
        # Zeige curl-äquivalenten Befehl an (nur zur Information)
        $curlEquivalent = @"
Entsprechender curl-Befehl:
curl $ApiUrl \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer $ApiKey" \
  -d '{
  "model": "$Model",
  "messages": [
    {
      "role": "system",
      "content": "Analysiere Windows-Ereignisdaten..."
    },
    {
      "role": "user", 
      "content": "[Ereignisdaten]"
    }
  ]
}'
"@
        Write-Host $curlEquivalent -ForegroundColor DarkGray

        # Tatsächliche Anfrage senden
        $response = Invoke-RestMethod -Uri $ApiUrl -Method Post -Headers $headers -Body $body -ErrorAction Stop
        
        # Hier nehmen wir an, dass die Antwort in $response.choices[0].message.content enthalten ist.
        Write-Host "Analyse erfolgreich empfangen." -ForegroundColor Green
        return $response.choices[0].message.content
    }
    catch {
        $errorMessage = "Fehler bei der API-Anfrage: $_"
        Write-Host $errorMessage -ForegroundColor Red
        
        # Detaillierte Fehlermeldung erstellen
        $detailedError = @"
# Fehler bei der Verbindung zu OpenRouter API

## Fehlermeldung
$_

## Moegliche Ursachen
- Ungültiger API-Schlüssel (überprüfe den Schlüssel in der .env Datei)
- Netzwerkprobleme oder keine Internetverbindung
- OpenRouter-Dienst ist möglicherweise nicht erreichbar
- Firewall oder Antivirus blockiert die Verbindung

## Überprüfung des API-Schlüssels
Aktuell verwendeter API-Schlüssel: `$(if ($ApiKey.Length -gt 8) { $ApiKey.Substring(0, 4) + "..." + $ApiKey.Substring($ApiKey.Length - 4) } else { "Ungültig oder zu kurz" })`
Konfigurationsdatei: `$configFile`

## Was du tun kannst
1. Überprüfe deine Internetverbindung
2. Stelle sicher, dass der API-Schlüssel in der .env Datei korrekt ist
3. Besuche [OpenRouter](https://openrouter.ai) um zu prüfen, ob der Dienst verfügbar ist

Du kannst auch die Demo-Version des Skripts ohne API ausprobieren: `.\DemoEventAnalyzer.ps1`
"@
        
        return $detailedError
    }
}

# --- Anfrage senden und Antwort empfangen ---
$analyseErgebnis = Get-AIAnalysis -LogData $logData -ApiUrl $apiUrl -ApiKey $apiKey -Model $aiModell

# --- GUI erstellen zur Darstellung der Analyse ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Event Analyzer - Claude 3.7 Sonnet"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.Icon = [System.Drawing.SystemIcons]::Information

# Menu erstellen
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Datei")
$saveMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Analyse speichern...")
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Beenden")

$saveMenuItem.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Textdatei (*.txt)|*.txt|Markdown (*.md)|*.md|All files (*.*)|*.*"
        $saveFileDialog.InitialDirectory = $outputDir
        $saveFileDialog.FileName = "Ereignisanalyse_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').md"
    
        if ($saveFileDialog.ShowDialog() -eq 'OK') {
            $textBox.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding utf8
            [System.Windows.Forms.MessageBox]::Show(
                "Analyse wurde gespeichert unter:`n$($saveFileDialog.FileName)", 
                "Gespeichert",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })

$exitMenuItem.Add_Click({
        $form.Close()
    })

$fileMenu.DropDownItems.Add($saveMenuItem)
$fileMenu.DropDownItems.Add($exitMenuItem)
$menuStrip.Items.Add($fileMenu)
$form.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip

# Panel für TextBox
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$panel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)

# TextBox erstellen für die Anzeige der Analyse
$textBox = New-Object System.Windows.Forms.RichTextBox
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$textBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$textBox.BackColor = [System.Drawing.Color]::White
$textBox.ForeColor = [System.Drawing.Color]::Black
$textBox.WordWrap = $true
$textBox.Text = $analyseErgebnis

$panel.Controls.Add($textBox)
$form.Controls.Add($panel)

# Status Bar hinzufügen
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ereignisprotokoll: $logName | Ereignisse: $maxEvents | KI-Modell: $aiModell"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Form anzeigen
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)

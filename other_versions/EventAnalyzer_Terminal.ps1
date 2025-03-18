# --- Windows Event Analyzer mit OpenRouter AI (Terminal-Version) ---
# Analysiert Windows-Ereignisprotokolle mit Hilfe von KI via OpenRouter API
# Terminal-Version mit nativer PowerShell-Farbunterstuetzung

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
    "Gemini flash 2.0"             = "google/gemini-2.0-flash-001"
}

# Standard-Modell ist Claude 3.7 Sonnet
$aiModell = $availableModels["Claude 3.7 Sonnet"]

# Klare Anzeige fuer Start des Programms
Clear-Host
Write-Host "Windows Event Analyzer (Terminal-Version)" -ForegroundColor Cyan -NoNewline
Write-Host " v1.0 " -ForegroundColor White -BackgroundColor Blue
Write-Host "Analysiert Ereignisse mit Hilfe von KI ueber die OpenRouter API" -ForegroundColor DarkGray
Write-Host "-".PadRight(80, "-")

# Modell-Auswahl anzeigen
Write-Host "`nVerfuegbare KI-Modelle:" -ForegroundColor Cyan
$i = 1
$modelList = @{}
foreach ($model in $availableModels.Keys) {
    $modelStr = "[$i] $model"
    if ($model -eq "Claude 3.7 Sonnet") {
        Write-Host $modelStr -ForegroundColor Green -NoNewline
        Write-Host " ($($availableModels[$model]))" -ForegroundColor DarkGray
    }
    else {
        Write-Host $modelStr -ForegroundColor Yellow -NoNewline
        Write-Host " ($($availableModels[$model]))" -ForegroundColor DarkGray
    }
    $modelList[$i] = $model
    $i++
}

$modelChoice = Read-Host "`nWaehle ein Modell (1-$($availableModels.Count)) [Standard: 1]"
if ($modelChoice -ne "" -and $modelChoice -match '^\d+$' -and [int]$modelChoice -ge 1 -and [int]$modelChoice -le $availableModels.Count) {
    $selectedModel = $modelList[[int]$modelChoice]
    $aiModell = $availableModels[$selectedModel]
    Write-Host "Ausgewaehltes Modell: " -ForegroundColor Green -NoNewline
    Write-Host "$selectedModel ($aiModell)" -ForegroundColor Green
}
else {
    $selectedModel = "Claude 3.7 Sonnet"
    Write-Host "Verwende Standard-Modell: " -ForegroundColor Green -NoNewline
    Write-Host "$selectedModel ($aiModell)" -ForegroundColor Green
}

# Ereignisanzahl festlegen
$maxEvents = 50
Write-Host "`nEreignisanzahl:" -ForegroundColor Cyan
$eventsChoice = Read-Host "Anzahl der zu analysierenden Ereignisse (50-250) [Standard: 50]"
if ($eventsChoice -ne "" -and $eventsChoice -match '^\d+$' -and [int]$eventsChoice -ge 50 -and [int]$eventsChoice -le 250) {
    $maxEvents = [int]$eventsChoice
}
Write-Host "Verwende Ereignisanzahl: $maxEvents" -ForegroundColor Green

# Test der Internetverbindung
Write-Host "`nTeste die Internetverbindung..." -ForegroundColor Yellow
try {
    $testConnection = Test-Connection -ComputerName "google.com" -Count 1 -Quiet
    if ($testConnection) {
        Write-Host "Internetverbindung OK" -ForegroundColor Green
    }
    else {
        Write-Host "Keine Internetverbindung verfuegbar. Die API kann nicht erreicht werden." -ForegroundColor Red
        Write-Host "Starte Demo-Modus stattdessen..." -ForegroundColor Yellow
        & "$PSScriptRoot\DemoEventAnalyzer.ps1"
        exit
    }
}
catch {
    Write-Host "Fehler beim Pruefen der Internetverbindung: $_" -ForegroundColor Red
    Write-Host "Starte Demo-Modus stattdessen..." -ForegroundColor Yellow
    & "$PSScriptRoot\DemoEventAnalyzer.ps1"
    exit
}

# API-Schluessel aus .env Datei laden oder abfragen
$apiKey = Get-EnvVariable -Key "OPENROUTER_API_KEY" -FilePath $configFile

if ($null -eq $apiKey -or $apiKey -eq "") {
    Write-Host "`n--- OpenRouter API-Schluessel Konfiguration ---" -ForegroundColor Cyan
    Write-Host "Ein OpenRouter API-Schluessel ist erforderlich fuer die Analyse mit $selectedModel."
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

# --- Funktion zum Erfassen der Ereignisdaten ---
function Get-EventLogData {
    param (
        [string]$LogName = "System",
        [int]$MaxEvents = 50
    )
    
    Write-Host "`nErfasse Ereignisdaten aus $LogName-Protokoll..." -ForegroundColor Yellow
    
    try {
        $events = Get-WinEvent -LogName $LogName -MaxEvents $MaxEvents -ErrorAction Stop | 
        Select-Object Id, LevelDisplayName, TimeCreated, Message
        
        Write-Host "$($events.Count) Ereignisse erfolgreich gesammelt." -ForegroundColor Green
        return $events | ConvertTo-Json -Depth 3
    }
    catch {
        Write-Error "Fehler beim Auslesen der Ereignisdaten: $_"
        return $null
    }
}

# --- Funktion zur API-Anfrage ---
function Get-AIAnalysis {
    param (
        [string]$LogData,
        [string]$ApiUrl,
        [string]$ApiKey,
        [string]$Model
    )
    
    Write-Host "`nSende Daten an $Model zur Analyse..." -ForegroundColor Yellow
    
    $systemPrompt = @"
Analysiere die folgenden Windows-Ereignisdaten und gib eine verstaendliche Zusammenfassung mit folgenden Abschnitten:
1. Uebersicht: Anzahl und Arten der Ereignisse
2. Wichtige Ereignisse: Hervorheben kritischer oder ungewoehnlicher Eintraege
3. Fehleranalyse: Moegliche Ursachen fuer Fehler oder Warnungen
4. Empfehlungen: Konkrete Handlungsempfehlungen basierend auf den Ereignissen
5. Zusammenfassung: Allgemeiner Systemzustand und wichtigste Punkte

Formatiere die Ausgabe mit Markdown fuer bessere Lesbarkeit.

WICHTIG: Verwende nur ASCII-Zeichen in deiner Antwort, um Encoding-Probleme zu vermeiden. 
Ersetze Umlaute wie folgt:
- 'ä' durch 'ae'
- 'ö' durch 'oe'
- 'ü' durch 'ue'
- 'Ä' durch 'Ae'
- 'Ö' durch 'Oe'
- 'Ü' durch 'Ue'
- 'ß' durch 'ss'
Vermeide alle sonstigen Sonderzeichen, die nicht im ASCII-Zeichensatz enthalten sind.
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
        # API-Anfrage mit Protokollierung
        Write-Host "`nSende API-Anfrage an OpenRouter:" -ForegroundColor Yellow
        Write-Host "URL: $ApiUrl" -ForegroundColor DarkGray
        Write-Host "Modell: $Model" -ForegroundColor DarkGray
        Write-Host "API-Schluessel: $(if ($ApiKey.Length -gt 8) { $ApiKey.Substring(0, 4) + "..." + $ApiKey.Substring($ApiKey.Length - 4) } else { "Ungueltig oder zu kurz" })" -ForegroundColor DarkGray
        
        # Fortschrittsanzeige starten
        Write-Host -NoNewline "Warte auf Antwort... " -ForegroundColor Yellow
        $progressTimer = [System.Diagnostics.Stopwatch]::StartNew()
        
        # Tatsaechliche Anfrage senden
        $response = Invoke-RestMethod -Uri $ApiUrl -Method Post -Headers $headers -Body $body -ErrorAction Stop
        
        # Hier nehmen wir an, dass die Antwort in $response.choices[0].message.content enthalten ist.
        $elapsedSeconds = [System.Math]::Round($progressTimer.Elapsed.TotalSeconds, 1)
        Write-Host "`rAnalyse erfolgreich empfangen in $elapsedSeconds Sekunden.                    " -ForegroundColor Green
        
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
- Ungueltiger API-Schluessel (ueberpruefe den Schluessel in der .env Datei)
- Netzwerkprobleme oder keine Internetverbindung
- OpenRouter-Dienst ist moeglicherweise nicht erreichbar
- Firewall oder Antivirus blockiert die Verbindung

## Ueberpruefung des API-Schluessels
Aktuell verwendeter API-Schluessel: `$(if ($ApiKey.Length -gt 8) { $ApiKey.Substring(0, 4) + "..." + $ApiKey.Substring($ApiKey.Length - 4) } else { "Ungueltig oder zu kurz" })`
Konfigurationsdatei: `$configFile`

## Was du tun kannst
1. Ueberpruefe deine Internetverbindung
2. Stelle sicher, dass der API-Schluessel in der .env Datei korrekt ist
3. Besuche [OpenRouter](https://openrouter.ai) um zu pruefen, ob der Dienst verfuegbar ist

Du kannst auch die Demo-Version des Skripts ohne API ausprobieren: `.\DemoEventAnalyzer.ps1`
"@
        
        return $detailedError
    }
}

# --- Funktion zum Parsen von Markdown fuer farbige Terminaldarstellung ---
function Format-MarkdownForTerminal {
    param (
        [string]$MarkdownText
    )
    
    # Zeilen aufteilen
    $lines = $MarkdownText -split "`r`n|\r|\n"
    $output = ""
    $inCodeBlock = $false
    
    # Durch jede Zeile gehen
    foreach ($line in $lines) {
        # Ueberschrift Level 1 (# Heading 1)
        if ($line -match '^# (.+)') {
            Write-Host $Matches[1] -ForegroundColor Cyan -BackgroundColor DarkBlue
            Write-Host ""
        }
        # Ueberschrift Level 2 (## Heading 2)
        elseif ($line -match '^## (.+)') {
            Write-Host $Matches[1] -ForegroundColor Cyan
            Write-Host ""
        }
        # Ueberschrift Level 3 (### Heading 3)
        elseif ($line -match '^### (.+)') {
            Write-Host $Matches[1] -ForegroundColor Blue
            Write-Host ""
        }
        # Aufzaehlungspunkte
        elseif ($line -match '^(\s*[-*+]\s+)(.+)$') {
            Write-Host "  * " -ForegroundColor Yellow -NoNewline
            Write-Host $Matches[2] -ForegroundColor White
        }
        # Code-Bloecke (nur Start/Ende Marker)
        elseif ($line -match '^```') {
            $inCodeBlock = !$inCodeBlock
            if ($inCodeBlock) {
                Write-Host ""
                Write-Host "+--- Code -------------------------------------------+" -ForegroundColor Magenta
            } 
            else {
                Write-Host "+---------------------------------------------------+" -ForegroundColor Magenta
                Write-Host ""
            }
        }
        # Inhalt von Code-Bloecken
        elseif ($inCodeBlock) {
            Write-Host "| $line" -ForegroundColor Magenta
        }
        # Horizontale Linie
        elseif ($line -match '^-{3,}$' -or $line -match '^_{3,}$' -or $line -match '^\*{3,}$') {
            Write-Host "---------------------------------------------------" -ForegroundColor DarkGray
        }
        # Links, fett, kursiv und normale Zeilen
        else {
            $formattedLine = $line
            
            # Links in einem einheitlichen Format darstellen: [Text](URL) -> Text (URL)
            $formattedLine = $formattedLine -replace '\[(.+?)\]\((.+?)\)', '$1 ($2)'
            
            # Normale Zeile
            if ($formattedLine.Trim() -ne "") {
                # Wenn die Zeile mit einem Bindestrich oder Sternchen beginnt, aber kein richtiger Listenpunkt ist
                if ($formattedLine -match '^\s*[-*+]') {
                    Write-Host $formattedLine -ForegroundColor Yellow
                }
                # Einfache Textzeile
                else {
                    Write-Host $formattedLine -ForegroundColor White
                }
            }
            else {
                Write-Host "" # Leerzeile
            }
        }
    }
}

# --- Ereignisdaten erfassen ---
$logData = Get-EventLogData -LogName "System" -MaxEvents $maxEvents
if ($null -eq $logData) {
    Write-Host "Fehler beim Erfassen der Ereignisdaten. Das Skript wird beendet." -ForegroundColor Red
    exit
}

# --- Anfrage senden und Antwort empfangen ---
$analyseErgebnis = Get-AIAnalysis -LogData $logData -ApiUrl $apiUrl -ApiKey $apiKey -Model $aiModell

# Ergebnis im Terminal anzeigen
Write-Host "`n ANALYSEERGEBNIS " -ForegroundColor White -BackgroundColor Blue
Write-Host "Modell: $selectedModel | Ereignisse: $maxEvents | Protokoll: System" -ForegroundColor DarkGray
Write-Host "-".PadRight(80, "-")

Format-MarkdownForTerminal -MarkdownText $analyseErgebnis

# Ergebnis speichern?
$saveChoice = Read-Host "`nMoechtest du das Ergebnis speichern? (j/n)"
if ($saveChoice -eq "j" -or $saveChoice -eq "J") {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $filename = "Ereignisanalyse_$timestamp.md"
    $filepath = Join-Path -Path $outputDir -ChildPath $filename
    
    try {
        # UTF-8 ohne BOM verwenden, um Encoding-Probleme zu vermeiden
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
        [System.IO.File]::WriteAllText($filepath, $analyseErgebnis, $Utf8NoBomEncoding)
        
        Write-Host "Analyse erfolgreich gespeichert unter:" -ForegroundColor Green
        Write-Host $filepath -ForegroundColor Cyan
    }
    catch {
        Write-Host "Fehler beim Speichern: $_" -ForegroundColor Red
    }
}

Write-Host "`nAnalyse abgeschlossen." -ForegroundColor Green
Write-Host "Druecke eine beliebige Taste zum Beenden..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

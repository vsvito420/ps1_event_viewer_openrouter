# --- Windows Event Analyzer (DEMO-VERSION) ---
# Analysiert Windows-Ereignisprotokolle und zeigt eine vordefinierte Analyse ohne API-Schluessel

# --- Konfiguration ---
# Setze die Ausgabecodierung auf UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Pfade fuer Speicherung
$appDir = "$env:USERPROFILE\Documents\EventAnalyzer"
$outputDir = $appDir

# Erstelle Verzeichnis falls nicht vorhanden
if (-not (Test-Path -Path $appDir)) {
    New-Item -ItemType Directory -Path $appDir | Out-Null
    Write-Host "Verzeichnis $appDir wurde erstellt." -ForegroundColor Green
}

Write-Host "`n--- DEMO-MODUS ---" -ForegroundColor Cyan
Write-Host "Diese Version benoetigt keinen API-Schluessel und zeigt vordefinierte Ergebnisse an."
Write-Host "Fuer die volle Funktionalitaet mit echter KI-Analyse, bitte EventAnalyzer.ps1 verwenden."
Write-Host "`nSammle Windows-Ereignisdaten (ohne API-Aufruf)..." -ForegroundColor Yellow

# --- Funktion zum Erfassen der Ereignisdaten ---
function Get-EventLogData {
    param (
        [string]$LogName = "System",
        [int]$MaxEvents = 10
    )
    
    try {
        $events = Get-WinEvent -LogName $LogName -MaxEvents $MaxEvents -ErrorAction Stop | 
        Select-Object Id, LevelDisplayName, TimeCreated, Message
        
        Write-Host "$($events.Count) Ereignisse erfolgreich gesammelt."
        return $events
    }
    catch {
        Write-Error "Fehler beim Auslesen der Ereignisdaten: $_"
        return $null
    }
}

# --- Ereignisdaten erfassen ---
$logName = "System"
$maxEvents = 10
$eventData = Get-EventLogData -LogName $logName -MaxEvents $maxEvents

if ($null -eq $eventData) {
    exit
}

# Anzeigen der Ereignistypen fuer die Demo
$eventStats = $eventData | Group-Object -Property LevelDisplayName | 
Select-Object Name, Count | 
Format-Table -AutoSize

Write-Host "`nGefundene Ereignistypen:" -ForegroundColor Cyan
$eventStats

# --- Simulierte KI-Analyse generieren ---
Write-Host "Erstelle Demo-Analyse..."

# Event-Typen zaehlen
$infoCount = ($eventData | Where-Object { $_.LevelDisplayName -eq "Information" }).Count
$warnCount = ($eventData | Where-Object { $_.LevelDisplayName -eq "Warning" }).Count
$errorCount = ($eventData | Where-Object { $_.LevelDisplayName -eq "Error" }).Count

# Simulierte API-Antwort mit tatsaechlichen Zahlen
$analyseErgebnis = @"
# Analyse der Windows-Ereignisdaten (DEMO)

## Uebersicht
- Insgesamt wurden $maxEvents Systemereignisse untersucht
- $infoCount Informationsmeldungen
- $warnCount Warnungen
- $errorCount Fehler

## Wichtige Ereignisse

### Windows-Dienste
Es wurden mehrere Ereignisse zu Windows-Diensten gefunden. Diese Ereignisse sind typischerweise Teil des normalen Systembetriebs und zeigen den Start, Stop oder die Aktualisierung von Diensten an.

### Systemstart-Ereignisse
Einige Ereignisse beziehen sich auf den Systemstart oder Energieverwaltung. Diese sind normal und treten bei jedem Systemstart auf.

## Fehleranalyse

### Moegliche Netzwerkprobleme
Es wurden Anzeichen fuer Netzwerkverbindungsschwankungen gefunden. Moeglicherweise gab es kurzfristige Verbindungsunterbrechungen zu lokalen Netzwerkressourcen oder dem Internet.

### Fehlgeschlagene Anmeldeversuche
In den Ereignisdaten wurden Hinweise auf fehlgeschlagene Anmeldeversuche gefunden. Dies kann normal sein, wenn Benutzer ihr Passwort falsch eingegeben haben, oder ein Hinweis auf Brute-Force-Versuche.

## Empfehlungen

1. **Systemupdates**: Stellen Sie sicher, dass das System auf dem neuesten Stand ist.
2. **Netzwerkdiagnose**: Bei anhaltenden Netzwerkproblemen sollten Sie eine Netzwerkdiagnose durchfuehren.
3. **Sicherheitsaudit**: Pruefen Sie die Anmeldeversuche regelmaessig, um unbefugte Zugriffsversuche zu erkennen.
4. **Dienste-Ueberwachung**: Richten Sie eine regelmaessige Ueberwachung wichtiger Systemdienste ein.

## Zusammenfassung

Das System befindet sich in einem normalen Betriebszustand. Die gefundenen Warnungen und Fehler sind nicht kritisch und weisen auf keine schwerwiegenden Probleme hin. Eine regelmaessige Wartung und Ueberwachung wird empfohlen, um die Systemstabilitaet zu gewaehrleisten.

---

**Hinweis**: Dies ist eine **DEMO-ANALYSE** und basiert auf einer vordefinierten Vorlage mit Einbeziehung der tatsaechlichen Ereignisanzahl. Fuer eine echte KI-Analyse mit Claude 3.7 Sonnet, verwenden Sie bitte die Hauptversion mit API-Schluessel.
"@

# --- Funktion zum Parsen von Markdown für farbige Darstellung ---
function Format-MarkdownText {
    param (
        [System.Windows.Forms.RichTextBox]$RichTextBox,
        [string]$MarkdownText
    )
    
    # Textbox leeren und zurücksetzen
    $RichTextBox.Clear()
    
    # Farben definieren
    $colorHeading1 = [System.Drawing.Color]::FromArgb(255, 77, 172, 253)    # Hellblau
    $colorHeading2 = [System.Drawing.Color]::FromArgb(255, 102, 204, 255)   # Blau
    $colorHeading3 = [System.Drawing.Color]::FromArgb(255, 129, 199, 247)   # Helleres Blau
    $colorBold = [System.Drawing.Color]::FromArgb(255, 255, 203, 107)       # Gelb-Orange
    $colorItalic = [System.Drawing.Color]::FromArgb(255, 180, 210, 115)     # Grün-Gelb
    $colorList = [System.Drawing.Color]::FromArgb(255, 247, 140, 108)       # Orange
    $colorCode = [System.Drawing.Color]::FromArgb(255, 190, 145, 255)       # Lila
    $colorNormal = [System.Drawing.Color]::FromArgb(255, 220, 220, 220)     # Hellgrau
    
    # Markdown-Text in Zeilen aufteilen
    $lines = $MarkdownText -split "`r`n|\r|\n"
    
    # Durch jede Zeile gehen
    foreach ($line in $lines) {
        $currentColor = $colorNormal
        $isBold = $false
        $isItalic = $false
        
        # Überschriften prüfen
        if ($line -match '^# (.+)') {
            $currentColor = $colorHeading1
            $line = $Matches[1]
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.FontFamily, 14, [System.Drawing.FontStyle]::Bold)
        }
        elseif ($line -match '^## (.+)') {
            $currentColor = $colorHeading2
            $line = $Matches[1]
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
        }
        elseif ($line -match '^### (.+)') {
            $currentColor = $colorHeading3
            $line = $Matches[1]
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.FontFamily, 11, [System.Drawing.FontStyle]::Bold)
        }
        # Aufzählungspunkte prüfen
        elseif ($line -match '^(\s*[-*+]\s+)(.+)$') {
            $prefix = $Matches[1]
            $content = $Matches[2]
            
            # Zuerst das Bullet-Point-Symbol hinzufügen
            $RichTextBox.SelectionColor = $colorList
            $RichTextBox.AppendText($prefix)
            
            # Dann den restlichen Inhalt
            $RichTextBox.SelectionColor = $colorNormal
            $RichTextBox.AppendText($content)
            $RichTextBox.AppendText("`n")
            continue
        }
        # Code-Blöcke prüfen
        elseif ($line -match '^```') {
            $currentColor = $colorCode
        }
        
        # Fett und kursiv für die gesamte Zeile prüfen - einfache Implementierung
        if ($line -match '\*\*(.+)\*\*') {
            $isBold = $true
            $line = $line -replace '\*\*(.+)\*\*', '$1'
        }
        if ($line -match '_(.+)_' -or $line -match '\*(.+)\*') {
            $isItalic = $true
            $line = $line -replace '_(.+)_', '$1'
            $line = $line -replace '\*(.+)\*', '$1'
        }
        
        # Schriftart anpassen
        if ($isBold -and $isItalic) {
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.FontFamily, $RichTextBox.Font.Size, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)
        }
        elseif ($isBold) {
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.FontFamily, $RichTextBox.Font.Size, [System.Drawing.FontStyle]::Bold)
        }
        elseif ($isItalic) {
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.FontFamily, $RichTextBox.Font.Size, [System.Drawing.FontStyle]::Italic)
        }
        else {
            $RichTextBox.SelectionFont = New-Object System.Drawing.Font($RichTextBox.Font.FontFamily, $RichTextBox.Font.Size, [System.Drawing.FontStyle]::Regular)
        }
        
        # Farbe setzen und Zeile hinzufügen
        $RichTextBox.SelectionColor = $currentColor
        $RichTextBox.AppendText($line + "`n")
    }
}

# --- GUI erstellen zur Darstellung der Analyse ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Dunkles Theme-Farben
$darkBackground = [System.Drawing.Color]::FromArgb(255, 30, 30, 30)
$darkMenuBackground = [System.Drawing.Color]::FromArgb(255, 45, 45, 45)
$darkText = [System.Drawing.Color]::FromArgb(255, 220, 220, 220)
$darkAccent = [System.Drawing.Color]::FromArgb(255, 0, 120, 215)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Event Analyzer - DEMO (Dunkles Theme)"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.Icon = [System.Drawing.SystemIcons]::Information
$form.BackColor = $darkBackground
$form.ForeColor = $darkText

# Menu erstellen mit dunklem Theme
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $darkMenuBackground
$menuStrip.ForeColor = $darkText
$menuStrip.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::Professional

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Datei")
$fileMenu.ForeColor = $darkText
$saveMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Analyse speichern...")
$saveMenuItem.ForeColor = $darkText
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Beenden")
$exitMenuItem.ForeColor = $darkText

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Hilfe")
$helpMenu.ForeColor = $darkText
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Ueber")
$aboutMenuItem.ForeColor = $darkText

$saveMenuItem.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Textdatei (*.txt)|*.txt|Markdown (*.md)|*.md|All files (*.*)|*.*"
        $saveFileDialog.InitialDirectory = $outputDir
        $saveFileDialog.FileName = "Ereignisanalyse_DEMO_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').md"
    
        if ($saveFileDialog.ShowDialog() -eq 'OK') {
            # UTF-8 ohne BOM verwenden, um Encoding-Probleme zu vermeiden
            $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
            [System.IO.File]::WriteAllText($saveFileDialog.FileName, $textBox.Text, $Utf8NoBomEncoding)
            
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

$aboutMenuItem.Add_Click({
        [System.Windows.Forms.MessageBox]::Show(
            "Windows Event Analyzer (DEMO-VERSION)`n`nDiese Demo-Version zeigt eine vordefinierte Analyse basierend auf tatsaechlichen Ereignisdaten. Fuer eine echte KI-Analyse, verwenden Sie bitte die Hauptversion mit API-Schluessel.", 
            "Ueber Windows Event Analyzer",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })

$fileMenu.DropDownItems.Add($saveMenuItem)
$fileMenu.DropDownItems.Add($exitMenuItem)
$helpMenu.DropDownItems.Add($aboutMenuItem)
$menuStrip.Items.Add($fileMenu)
$menuStrip.Items.Add($helpMenu)
$form.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip

# Panel für TextBox
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$panel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$panel.BackColor = $darkBackground

# RichTextBox erstellen für die Anzeige der Analyse
$textBox = New-Object System.Windows.Forms.RichTextBox
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$textBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$textBox.BackColor = $darkBackground
$textBox.ForeColor = $darkText
$textBox.WordWrap = $true

# Text mit Markdown-Formatierung setzen
Format-MarkdownText -RichTextBox $textBox -MarkdownText $analyseErgebnis

$panel.Controls.Add($textBox)
$form.Controls.Add($panel)

# Status Bar hinzufügen mit dunklem Theme
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $darkMenuBackground
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "DEMO-MODUS | Ereignisprotokoll: $logName | Ereignisse: $maxEvents | Keine API-Verbindung erforderlich"
$statusLabel.ForeColor = $darkText
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

Write-Host "`nStarte die Demo-GUI..." -ForegroundColor Green
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)

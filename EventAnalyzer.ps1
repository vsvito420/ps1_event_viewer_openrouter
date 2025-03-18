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
    "Gemini flash 2.0"             = "google/gemini-2.0-flash-001"
}

# Standard-Modell ist Claude 3.7 Sonnet
$aiModell = $availableModels["Claude 3.7 Sonnet"]

# --- GUI erstellen zur Darstellung der Analyse ---
# Hinweis: GUI wird jetzt VOR der Datenerfassung erstellt
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Dunkles Theme-Farben
$darkBackground = [System.Drawing.Color]::FromArgb(255, 30, 30, 30)
$darkMenuBackground = [System.Drawing.Color]::FromArgb(255, 45, 45, 45)
$darkText = [System.Drawing.Color]::FromArgb(255, 220, 220, 220)
$darkAccent = [System.Drawing.Color]::FromArgb(255, 0, 120, 215)
$darkControlBackground = [System.Drawing.Color]::FromArgb(255, 40, 40, 40)

# Formular erstellen - Position angepasst, um Inhaltsabschneiden zu vermeiden
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Event Analyzer - Dunkles Theme"
$form.Size = New-Object System.Drawing.Size(1100, 700)
$form.StartPosition = "CenterScreen"  # Automatisch zentrieren
$form.Icon = [System.Drawing.SystemIcons]::Information
$form.BackColor = $darkBackground
$form.ForeColor = $darkText
$form.MinimumSize = New-Object System.Drawing.Size(800, 500)  # Minimalgröße hinzugefügt

# Import für DataGridView-Styling
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Globale Variablen für die Ansichtsmodi
$script:viewMode = "both"  # Mögliche Werte: "both", "text", "grid"

# Seitenleiste für Menü und Konfiguration
$sidebarPanel = New-Object System.Windows.Forms.Panel
$sidebarPanel.Dock = [System.Windows.Forms.DockStyle]::Left
$sidebarPanel.Width = 220
$sidebarPanel.BackColor = $darkControlBackground
$sidebarPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)

# Funktion zum Umschalten des Anzeigemodus
function Set-ViewMode {
    param(
        [string]$Mode # "both", "text", "grid"
    )

    $script:viewMode = $Mode
    
    switch ($Mode) {
        "both" {
            # Beide Panels sichtbar - Text links, Tabelle rechts
            $splitContainer.Orientation = [System.Windows.Forms.Orientation]::Vertical
            $splitContainer.SplitterDistance = [Math]::Min([Math]::Max(550, $splitContainer.Width / 2), $splitContainer.Width * 0.6)
            $splitContainer.Panel1Collapsed = $false
            $splitContainer.Panel2Collapsed = $false
            $bothViewButton.BackColor = $darkAccent
            $textViewButton.BackColor = $darkMenuBackground
            $gridViewButton.BackColor = $darkMenuBackground
        }
        "text" {
            # Nur Textpanel anzeigen (links)
            $splitContainer.Panel1Collapsed = $false
            $splitContainer.Panel2Collapsed = $true
            $bothViewButton.BackColor = $darkMenuBackground
            $textViewButton.BackColor = $darkAccent
            $gridViewButton.BackColor = $darkMenuBackground
        }
        "grid" {
            # Nur Tabelle anzeigen (rechts)
            $splitContainer.Panel1Collapsed = $true
            $splitContainer.Panel2Collapsed = $false
            $bothViewButton.BackColor = $darkMenuBackground
            $textViewButton.BackColor = $darkMenuBackground
            $gridViewButton.BackColor = $darkAccent
        }
    }
}

# Container für Konfigurationselemente in der Sidebar
$configPanel = New-Object System.Windows.Forms.Panel
$configPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$configPanel.Height = 320
$configPanel.BackColor = $darkControlBackground
$configPanel.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)

# Überschrift für die Seitenleiste
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Windows Event Analyzer"
$titleLabel.ForeColor = $darkText
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(200, 25)
$sidebarPanel.Controls.Add($titleLabel)

# Trennlinie unter der Überschrift
$separatorLabel1 = New-Object System.Windows.Forms.Label
$separatorLabel1.Text = ""
$separatorLabel1.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$separatorLabel1.Location = New-Object System.Drawing.Point(10, 40)
$separatorLabel1.Size = New-Object System.Drawing.Size(200, 2)
$sidebarPanel.Controls.Add($separatorLabel1)

# Einstellungen in der Sidebar
$modelLabel = New-Object System.Windows.Forms.Label
$modelLabel.Text = "KI-Modell:"
$modelLabel.ForeColor = $darkText
$modelLabel.Location = New-Object System.Drawing.Point(10, 50)
$modelLabel.Size = New-Object System.Drawing.Size(80, 20)
$sidebarPanel.Controls.Add($modelLabel)

$modelComboBox = New-Object System.Windows.Forms.ComboBox
$modelComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$modelComboBox.BackColor = $darkBackground
$modelComboBox.ForeColor = $darkText
$modelComboBox.Location = New-Object System.Drawing.Point(10, 70)
$modelComboBox.Size = New-Object System.Drawing.Size(200, 20)
$modelComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

foreach ($model in $availableModels.Keys) {
    [void]$modelComboBox.Items.Add($model)
}
$modelComboBox.SelectedItem = "Claude 3.7 Sonnet"  # Standard-Modell auswählen
$sidebarPanel.Controls.Add($modelComboBox)

# Ereignisanzahl Slider
$eventsLabel = New-Object System.Windows.Forms.Label
$eventsLabel.Text = "Ereignisanzahl: 50"
$eventsLabel.ForeColor = $darkText
$eventsLabel.Location = New-Object System.Drawing.Point(10, 100)
$eventsLabel.Size = New-Object System.Drawing.Size(150, 20)
$sidebarPanel.Controls.Add($eventsLabel)

# TrackBar für Ereignisanzahl (50-250)
$eventsSlider = New-Object System.Windows.Forms.TrackBar
$eventsSlider.Minimum = 50
$eventsSlider.Maximum = 250
$eventsSlider.Value = 50  # Standardwert
$eventsSlider.TickFrequency = 25
$eventsSlider.LargeChange = 25
$eventsSlider.SmallChange = 5
$eventsSlider.Location = New-Object System.Drawing.Point(10, 120)
$eventsSlider.Size = New-Object System.Drawing.Size(200, 45)
$eventsSlider.BackColor = $darkControlBackground
$sidebarPanel.Controls.Add($eventsSlider)

# Label für aktuellen Wert
$eventsSlider.Add_ValueChanged({
        $eventsLabel.Text = "Ereignisanzahl: $($eventsSlider.Value)"
    })

# Notizbox für zusätzliche Anweisungen an die KI
$notizLabel = New-Object System.Windows.Forms.Label
$notizLabel.Text = "Zusätzliche Anweisungen:"
$notizLabel.ForeColor = $darkText
$notizLabel.Location = New-Object System.Drawing.Point(10, 165)
$notizLabel.Size = New-Object System.Drawing.Size(150, 20)
$sidebarPanel.Controls.Add($notizLabel)

$notizTextBox = New-Object System.Windows.Forms.TextBox
$notizTextBox.Location = New-Object System.Drawing.Point(10, 185)
$notizTextBox.Size = New-Object System.Drawing.Size(200, 60)
$notizTextBox.Multiline = $true
$notizTextBox.BackColor = $darkBackground
$notizTextBox.ForeColor = $darkText
$notizTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$notizTextBox.Text = "z.B. Ignoriere Programme wie Chrome oder Outlook"
$sidebarPanel.Controls.Add($notizTextBox)

# Trennlinie vor Ansichtsmodus
$separatorLabel2 = New-Object System.Windows.Forms.Label
$separatorLabel2.Text = ""
$separatorLabel2.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$separatorLabel2.Location = New-Object System.Drawing.Point(10, 255)
$separatorLabel2.Size = New-Object System.Drawing.Size(200, 2)
$sidebarPanel.Controls.Add($separatorLabel2)

# Überschrift für Ansichtsmodus
$viewModeLabel = New-Object System.Windows.Forms.Label
$viewModeLabel.Text = "Ansichtsmodus:"
$viewModeLabel.ForeColor = $darkText
$viewModeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$viewModeLabel.Location = New-Object System.Drawing.Point(10, 265)
$viewModeLabel.Size = New-Object System.Drawing.Size(150, 20)
$sidebarPanel.Controls.Add($viewModeLabel)

# Buttons für Ansichtsmodus
$bothViewButton = New-Object System.Windows.Forms.Button
$bothViewButton.Text = "Text und Tabelle"
$bothViewButton.BackColor = $darkAccent  # Standardmäßig aktiv
$bothViewButton.ForeColor = $darkText
$bothViewButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$bothViewButton.Location = New-Object System.Drawing.Point(10, 290)
$bothViewButton.Size = New-Object System.Drawing.Size(200, 30)
$bothViewButton.Add_Click({ Set-ViewMode -Mode "both" })
$sidebarPanel.Controls.Add($bothViewButton)

$textViewButton = New-Object System.Windows.Forms.Button
$textViewButton.Text = "Nur Text"
$textViewButton.BackColor = $darkMenuBackground
$textViewButton.ForeColor = $darkText
$textViewButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$textViewButton.Location = New-Object System.Drawing.Point(10, 325)
$textViewButton.Size = New-Object System.Drawing.Size(200, 30)
$textViewButton.Add_Click({ Set-ViewMode -Mode "text" })
$sidebarPanel.Controls.Add($textViewButton)

$gridViewButton = New-Object System.Windows.Forms.Button
$gridViewButton.Text = "Nur Tabelle"
$gridViewButton.BackColor = $darkMenuBackground
$gridViewButton.ForeColor = $darkText
$gridViewButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$gridViewButton.Location = New-Object System.Drawing.Point(10, 360)
$gridViewButton.Size = New-Object System.Drawing.Size(200, 30)
$gridViewButton.Add_Click({ Set-ViewMode -Mode "grid" })
$sidebarPanel.Controls.Add($gridViewButton)

# Trennlinie vor Aktionen
$separatorLabel3 = New-Object System.Windows.Forms.Label
$separatorLabel3.Text = ""
$separatorLabel3.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$separatorLabel3.Location = New-Object System.Drawing.Point(10, 400)
$separatorLabel3.Size = New-Object System.Drawing.Size(200, 2)
$sidebarPanel.Controls.Add($separatorLabel3)

# Überschrift für Aktionen
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = "Aktionen:"
$actionLabel.ForeColor = $darkText
$actionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$actionLabel.Location = New-Object System.Drawing.Point(10, 410)
$actionLabel.Size = New-Object System.Drawing.Size(100, 20)
$sidebarPanel.Controls.Add($actionLabel)

# Analyse-Button hinzufügen
$analyzeButton = New-Object System.Windows.Forms.Button
$analyzeButton.Text = "Ereignisse analysieren"
$analyzeButton.BackColor = $darkAccent
$analyzeButton.ForeColor = $darkText
$analyzeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$analyzeButton.Location = New-Object System.Drawing.Point(10, 435)
$analyzeButton.Size = New-Object System.Drawing.Size(200, 40)
$analyzeButton.Add_Click({ PerformAnalysis })
$sidebarPanel.Controls.Add($analyzeButton)

# Speichern-Button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Analyse speichern..."
$saveButton.BackColor = $darkMenuBackground
$saveButton.ForeColor = $darkText
$saveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$saveButton.Location = New-Object System.Drawing.Point(10, 480)
$saveButton.Size = New-Object System.Drawing.Size(200, 30)
$saveButton.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Textdatei (*.txt)|*.txt|Markdown (*.md)|*.md|All files (*.*)|*.*"
        $saveFileDialog.InitialDirectory = $outputDir
        $saveFileDialog.FileName = "Ereignisanalyse_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').md"

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
$sidebarPanel.Controls.Add($saveButton)

# Beenden-Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Beenden"
$exitButton.BackColor = $darkMenuBackground
$exitButton.ForeColor = $darkText
$exitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$exitButton.Location = New-Object System.Drawing.Point(10, 515)
$exitButton.Size = New-Object System.Drawing.Size(200, 30)
$exitButton.Add_Click({ $form.Close() })
$sidebarPanel.Controls.Add($exitButton)

# SplitContainer erstellen für geteilte Ansicht (links Text, rechts Tabelle)
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Vertical  # Vertical = links/rechts
$splitContainer.SplitterDistance = 550  # Anfängliche Teilung (links breiter)
$splitContainer.BackColor = $darkBackground
$splitContainer.Panel1.BackColor = $darkBackground
$splitContainer.Panel1.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$splitContainer.Panel2.BackColor = $darkBackground
$splitContainer.Panel2.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)

# RichTextBox für die linke Seite (Analyse)
$textBox = New-Object System.Windows.Forms.RichTextBox
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$textBox.Font = New-Object System.Drawing.Font("Consolas", 12)  # Größere Schrift (war 10)
$textBox.BackColor = $darkBackground
$textBox.ForeColor = $darkText
$textBox.WordWrap = $true
$textBox.Text = "Willkommen zum Windows Event Analyzer!`n`nWähle ein KI-Modell und die Anzahl der Ereignisse aus, und klicke dann auf 'Ereignisse analysieren', um die Analyse zu starten."

$splitContainer.Panel1.Controls.Add($textBox)

# DataGridView für die rechte Seite (Ereignistabelle)
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill
$dataGridView.BackgroundColor = $darkBackground
$dataGridView.ForeColor = [System.Drawing.Color]::Black
$dataGridView.GridColor = $darkControlBackground
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$dataGridView.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::Single
$dataGridView.RowHeadersVisible = $false
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.AllowUserToResizeRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = $darkMenuBackground
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = $darkText
$dataGridView.DefaultCellStyle.BackColor = $darkControlBackground
$dataGridView.DefaultCellStyle.ForeColor = $darkText
$dataGridView.DefaultCellStyle.SelectionBackColor = $darkAccent
$dataGridView.DefaultCellStyle.SelectionForeColor = $darkText
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersHeight = 36  # Erhöht (war 30)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

# Größere Schrift für DataGridView
$dataGridView.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

# Initialisieren mit leeren Spalten für die Analysetabelle
$columns = @(
    @{Name = "Problem"; Header = "Problem"; Width = 150 }, # Prägnanter Name statt allgemeiner Kategorie
    @{Name = "Beschreibung"; Header = "Beschreibung"; Width = 200 },
    @{Name = "Haeufigkeit"; Header = "Häufigkeit"; Width = 80 },
    @{Name = "Wichtigkeit"; Header = "Wichtigkeit"; Width = 80 },
    @{Name = "Fehlerbehebung"; Header = "Fehlerbehebung"; Width = 200 }
)

foreach ($column in $columns) {
    $newColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $newColumn.Name = $column.Name
    $newColumn.HeaderText = $column.Header
    $newColumn.Width = $column.Width
    $dataGridView.Columns.Add($newColumn)
}

$splitContainer.Panel2.Controls.Add($dataGridView)

# Panel für den SplitContainer
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$panel.BackColor = $darkBackground
$panel.Controls.Add($splitContainer)

# Status Bar hinzufügen
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $darkMenuBackground
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Bereit | Ereignisprotokoll: System | Modell: $aiModell"
$statusLabel.ForeColor = $darkText
$statusStrip.Items.Add($statusLabel)

# Die Menüleiste wurde entfernt und durch die Seitenleiste ersetzt
# Alle Menübefehle werden nun direkt über die Seitenleiste ausgeführt

# Elemente zum Formular hinzufügen
$form.Controls.Add($statusStrip)
$form.Controls.Add($panel)
$form.Controls.Add($sidebarPanel)

# --- Funktion zum Erfassen der Ereignisdaten ---
function Get-EventLogData {
    param (
        [string]$LogName = "System",
        [int]$MaxEvents = 50
    )
    
    Write-Host "Erfasse Ereignisdaten aus $LogName-Protokoll..."
    $statusLabel.Text = "Erfasse Ereignisdaten aus $LogName-Protokoll..."
    $form.Refresh()
    
    try {
        $global:eventData = Get-WinEvent -LogName $LogName -MaxEvents $MaxEvents -ErrorAction Stop | 
        Select-Object Id, LevelDisplayName, TimeCreated, Message
        
        Write-Host "$($global:eventData.Count) Ereignisse erfolgreich gesammelt."
        $statusLabel.Text = "$($global:eventData.Count) Ereignisse erfolgreich gesammelt."
        $form.Refresh()
        
        # DataGridView leeren - hier keine Anzeige der Rohdaten mehr
        $dataGridView.Rows.Clear()
        
        return $global:eventData | ConvertTo-Json -Depth 3
    }
    catch {
        Write-Error "Fehler beim Auslesen der Ereignisdaten: $_"
        $statusLabel.Text = "Fehler beim Auslesen der Ereignisdaten"
        $form.Refresh()
        
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
    
    Write-Host "Sende Daten an $Model zur Analyse..."
    $statusLabel.Text = "Sende Daten an $Model zur Analyse..."
    $form.Refresh()
    
    # Benutzernotizen für die KI verwenden, falls vorhanden
    $userNotes = $notizTextBox.Text.Trim()
    $userNotesText = ""
    if ($userNotes -ne "" -and $userNotes -ne "z.B. Ignoriere Programme wie Chrome oder Outlook") {
        $userNotesText = "`nZUSAETZLICHE BENUTZERANWEISUNGEN: $userNotes"
    }
    
    $systemPrompt = @"
Analysiere die folgenden Windows-Ereignisdaten und liefere das Ergebnis in zwei Teilen:$userNotesText

TEIL 1: MARKDOWN-ANALYSE
Erstelle eine verstaendliche Zusammenfassung mit folgenden Abschnitten:
- Uebersicht: Anzahl und Arten der Ereignisse
- Wichtige Ereignisse: Hervorheben kritischer oder ungewoehnlicher Eintraege
- Fehleranalyse: Moegliche Ursachen fuer Fehler oder Warnungen
- Empfehlungen: Konkrete Handlungsempfehlungen basierend auf den Ereignissen
- Zusammenfassung: Allgemeiner Systemzustand und wichtigste Punkte

Formatiere diesen Teil mit Markdown fuer bessere Lesbarkeit.

TEIL 2: TABELLENDATEN
Nach der Markdown-Analyse fuege einen JSON-Block ein, der eine tabellarische Zusammenfassung der Analyse enthaelt. 
Die Tabelle sollte folgende Struktur haben:

```json
{
  "table_rows": [
    {
      "kategorie": "Praeznanter Problemname", 
      "beschreibung": "Kurze Erklaerung des Problems",
      "haeufigkeit": "Anzahl oder Prozent", 
      "wichtigkeit": "1-10",
      "fehlerbehebung": "Konkrete Loesungsempfehlung"
    }
  ]
}
```

Fuer das "kategorie"-Feld sollst du praegnante, kurze Namen verwenden, die das Problem auf den Punkt bringen,
wie z.B. "Netzwerkausfall", "Speicherknappheit", "Treiberfehler", "Windows Update Problem", usw.

Das Feld "fehlerbehebung" muss konkrete, praxisnahe und spezifische Loesungsvorschlaege enthalten,
wie z.B. "Dienst xyz neu starten", "Treiber aktualisieren", "Registry-Schluessel anpassen", etc.
Diese sollten moeglichst direkt anwendbar und verstaendlich sein.

Die Tabelleneintraege sollten wichtige Kategorien aus deiner Analyse darstellen, wie z.B.:
- Haeufigste Ereignistypen
- Kritische Fehler
- Dienstwarnungen
- Systemprobleme
- Sicherheitshinweise
- Ressourcenengpaesse, etc.

Sortiere die Eintraege nach Wichtigkeit (1-10, wobei 10 am wichtigsten ist).

WICHTIG: Verwende nur ASCII-Zeichen in deiner Antwort, um Encoding-Probleme zu vermeiden. 
Ersetze Umlaute wie folgt:
- 'ae' statt 'ä'
- 'oe' statt 'ö'
- 'ue' statt 'ü'
- 'Ae' statt 'Ä'
- 'Oe' statt 'Ö'
- 'Ue' statt 'Ü'
- 'ss' statt 'ß'
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
        $statusLabel.Text = "Analyse erfolgreich empfangen | Ereignisprotokoll: System | Ereignisse: $($eventsSlider.Value) | Modell: $Model"
        $form.Refresh()
        
        return $response.choices[0].message.content
    }
    catch {
        $errorMessage = "Fehler bei der API-Anfrage: $_"
        Write-Host $errorMessage -ForegroundColor Red
        $statusLabel.Text = "Fehler bei der API-Anfrage"
        $form.Refresh()
        
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

# Funktion zur Durchführung der Analyse
function PerformAnalysis {
    # Aktuelle Einstellungen abrufen
    $maxEvents = $eventsSlider.Value
    $selectedModelName = $modelComboBox.SelectedItem
    $aiModell = $availableModels[$selectedModelName]
    $logName = "System"  # Hier könnte man später ein Dropdown für verschiedene Logs hinzufügen
    
    # Internetverbindung testen
    $textBox.Clear()
    $textBox.AppendText("Teste Internetverbindung...\n")
    $form.Refresh()
    
    try {
        $testConnection = Test-Connection -ComputerName "google.com" -Count 1 -Quiet
        if (-not $testConnection) {
            $textBox.AppendText("Keine Internetverbindung verfügbar. Die API kann nicht erreicht werden.\n")
            $textBox.AppendText("Starte Demo-Modus stattdessen...\n")
            $form.Refresh()
            
            # Start DemoEventAnalyzer
            & "$PSScriptRoot\DemoEventAnalyzer.ps1"
            $form.Close()
            return
        }
    }
    catch {
        $textBox.AppendText("Fehler beim Prüfen der Internetverbindung: $_\n")
        $textBox.AppendText("Starte Demo-Modus stattdessen...\n")
        $form.Refresh()
        
        # Start DemoEventAnalyzer
        & "$PSScriptRoot\DemoEventAnalyzer.ps1"
        $form.Close()
        return
    }
    
    # API-Schlüssel überprüfen
    $apiKey = Get-EnvVariable -Key "OPENROUTER_API_KEY" -FilePath $configFile
    
    if ($null -eq $apiKey -or $apiKey -eq "") {
        $textBox.Clear()
        $textBox.AppendText("Ein OpenRouter API-Schlüssel ist erforderlich.\n")
        $form.Refresh()
        
        # API-Schlüssel-Abfrage mit TextBox
        $apiKeyForm = New-Object System.Windows.Forms.Form
        $apiKeyForm.Text = "API-Schlüssel eingeben"
        $apiKeyForm.Size = New-Object System.Drawing.Size(400, 150)
        $apiKeyForm.StartPosition = "CenterParent"
        $apiKeyForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $apiKeyForm.BackColor = $darkBackground
        $apiKeyForm.ForeColor = $darkText
        
        $apiKeyLabel = New-Object System.Windows.Forms.Label
        $apiKeyLabel.Text = "Bitte gib deinen OpenRouter API-Schlüssel ein:"
        $apiKeyLabel.ForeColor = $darkText
        $apiKeyLabel.Location = New-Object System.Drawing.Point(10, 20)
        $apiKeyLabel.Size = New-Object System.Drawing.Size(380, 20)
        
        $apiKeyTextBox = New-Object System.Windows.Forms.TextBox
        $apiKeyTextBox.Location = New-Object System.Drawing.Point(10, 50)
        $apiKeyTextBox.Size = New-Object System.Drawing.Size(365, 20)
        $apiKeyTextBox.BackColor = $darkBackground
        $apiKeyTextBox.ForeColor = $darkText
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $okButton.BackColor = $darkAccent
        $okButton.ForeColor = $darkText
        $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $okButton.Location = New-Object System.Drawing.Point(275, 80)
        $okButton.Size = New-Object System.Drawing.Size(100, 30)
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Abbrechen"
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $cancelButton.BackColor = $darkMenuBackground
        $cancelButton.ForeColor = $darkText
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Location = New-Object System.Drawing.Point(165, 80)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
        
        $apiKeyForm.Controls.Add($apiKeyLabel)
        $apiKeyForm.Controls.Add($apiKeyTextBox)
        $apiKeyForm.Controls.Add($okButton)
        $apiKeyForm.Controls.Add($cancelButton)
        $apiKeyForm.AcceptButton = $okButton
        $apiKeyForm.CancelButton = $cancelButton
        
        $result = $apiKeyForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $apiKey = $apiKeyTextBox.Text.Trim()
            if ($apiKey -ne "") {
                Set-EnvVariable -Key "OPENROUTER_API_KEY" -Value $apiKey -FilePath $configFile
                Write-Host "API-Schlüssel erfolgreich gespeichert." -ForegroundColor Green
            }
            else {
                $textBox.AppendText("Kein API-Schlüssel eingegeben. Die Analyse kann nicht durchgeführt werden.\n")
                return
            }
        }
        else {
            $textBox.AppendText("API-Schlüssel-Eingabe abgebrochen. Die Analyse kann nicht durchgeführt werden.\n")
            return
        }
    }
    
    # Ereignisdaten laden
    $textBox.AppendText("Erfasse Ereignisdaten...\n")
    $form.Refresh()
    
    $logData = Get-EventLogData -LogName $logName -MaxEvents $maxEvents
    if ($null -eq $logData) {
        $textBox.AppendText("Fehler beim Erfassen der Ereignisdaten.\n")
        return
    }
    
    # API-Anfrage senden
    $textBox.AppendText("Sende Anfrage an KI-Modell...\n")
    $form.Refresh()
    
    $analyseErgebnis = Get-AIAnalysis -LogData $logData -ApiUrl $apiUrl -ApiKey $apiKey -Model $aiModell
    
    # Prüfen, ob JSON-Daten im Ergebnis enthalten sind
    if ($analyseErgebnis -match '```json\s*({[\s\S]*?})\s*```') {
        $jsonPart = $Matches[1]
        try {
            # Versuche, den JSON-Teil zu parsen
            $tableData = $jsonPart | ConvertFrom-Json
            
            # DataGridView leeren
            $dataGridView.Rows.Clear()
            
            # Tabellenzeilen aus der AI-Analyse in das DataGridView einfügen
            if ($tableData.table_rows -and $tableData.table_rows.Count -gt 0) {
                foreach ($row in $tableData.table_rows) {
                    # Daten für die Tabelle vorbereiten
                    $kategorie = $row.kategorie
                    $beschreibung = $row.beschreibung
                    $haeufigkeit = $row.haeufigkeit
                    $wichtigkeit = $row.wichtigkeit
                    $fehlerbehebung = $row.fehlerbehebung
                    
                    # In DataGridView einfügen
                    $rowIndex = $dataGridView.Rows.Add($kategorie, $beschreibung, $haeufigkeit, $wichtigkeit, $fehlerbehebung)
                    
                    # Wichtigkeit als Zellfarbe darstellen (je höher, desto intensiver)
                    # Konvertieren zu Integer, falls es als String kommt
                    try {
                        if ($wichtigkeit -match "^\d+-\d+$") {
                            # Falls Format "1-10" ist, nehme den höheren Wert
                            $wichtigkeit = [int]($wichtigkeit.Split('-')[1])
                        }
                        else {
                            $wichtigkeit = [int]$wichtigkeit
                        }
                    }
                    catch {
                        # Fallback, falls Konvertierung fehlschlägt
                        $wichtigkeit = 5
                    }
                    $priority = [Math]::Min([Math]::Max($wichtigkeit, 1), 10)  # Zwischen 1-10 begrenzen
                    
                    # Farbintensität basierend auf Wichtigkeit
                    $colorIntensity = 80 + ($priority * 15)  # 80-230 Bereich
                    
                    # Farbcodierung nach Wichtigkeit
                    if ($priority -ge 8) {
                        # Hohe Wichtigkeit (8-10): Rot
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, $colorIntensity, 40, 40)
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 200, 200)
                    }
                    elseif ($priority -ge 5) {
                        # Mittlere Wichtigkeit (5-7): Gelb/Orange
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, $colorIntensity, [Math]::Min($colorIntensity - 30, 200), 40)
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 240, 180)
                    }
                    elseif ($priority -ge 3) {
                        # Niedrigere Wichtigkeit (3-4): Blau
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 40, 40, [Math]::Min($colorIntensity, 230))
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.ForeColor = $darkText
                    }
                    else {
                        # Niedrigste Wichtigkeit (1-2): Grün
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 40, [Math]::Min($colorIntensity, 230), 40)
                        $dataGridView.Rows[$rowIndex].DefaultCellStyle.ForeColor = $darkText
                    }
                }
                
                $statusLabel.Text = "Analyse abgeschlossen | $($tableData.table_rows.Count) Eintraege in der Tabelle gefunden"
            }
        }
        catch {
            Write-Host "Fehler beim Parsen des JSON-Teils: $_" -ForegroundColor Red
        }
    }
    
    # Einfachen Text (ohne JSON-Teile) aus der Antwort extrahieren und formatieren
    $markdownPart = $analyseErgebnis -replace '```json[\s\S]*?```', ''
    Format-MarkdownText -RichTextBox $textBox -MarkdownText $markdownPart
}

# Analyse-Button-Handler
$analyzeButton.Add_Click({
        PerformAnalysis
    })

# Initial View Mode setzen
$form.Add_Shown({
        $form.Activate()
        # Standard-Ansichtsmodus "both" anwenden
        Set-ViewMode -Mode "both"
    })
[System.Windows.Forms.Application]::Run($form)

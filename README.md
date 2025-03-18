# Windows Event Analyzer mit OpenRouter AI

Dieses PowerShell-Skript analysiert Windows-Ereignisprotokolle mit Hilfe von Claude 3.7 Sonnet via OpenRouter.

## Direktausführung

Um das Skript direkt ohne Download auszuführen, kopieren Sie den folgenden Befehl in eine PowerShell-Konsole (als Administrator):

```powershell
powershell -Command "iwr -useb https://raw.githubusercontent.com/vsvito420/ps1_event_viewer_openrouter/main/EventAnalyzer.ps1 | iex"
```

## Features

- Analyse von Windows-Ereignisprotokollen mit KI
- Unterstützung für verschiedene KI-Modelle über OpenRouter
- Moderne GUI mit dunklem Theme
- Detaillierte Markdown-formatierte Analysen
- Tabellarische Übersicht der erkannten Probleme

## Installation

1. Laden Sie das Skript herunter oder nutzen Sie den One-Liner oben
2. Sie benötigen einen OpenRouter API-Schlüssel (https://openrouter.ai)
3. Bei der ersten Ausführung werden Sie nach dem API-Schlüssel gefragt

## Anforderungen

- Windows 10/11
- PowerShell 5.1 oder höher
- Internetverbindung
- OpenRouter API-Schlüssel

## Lizenz

[MIT License](LICENSE)

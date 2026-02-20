# MosaicTools Project

## For New Claude Sessions

This file is loaded at the start of every Claude Code session. If you're a new session:

1. **Don't re-analyze** — the architecture section below has the full codebase structure
2. **Check recent commits** — `git log --oneline -10`
3. **Read the relevant file** — paths are in the Architecture section
4. **Build command is below** — just run it, don't ask for confirmation

If making significant structural changes, update the Architecture section so future sessions stay current.

---

## ClarioIgnore - Separate Project

**IMPORTANT:** `ClarioIgnore/` is a separate tool. NOT part of MosaicTools releases — never include it in GitHub releases.

---

## Build Instructions

**When user says "build" or "rebuild", OR when you finish making code changes, run the full build command immediately without asking.**

### Key Paths
- **dotnet SDK** (NOT in PATH): `c:\Users\erik.richter\Desktop\dotnet\dotnet.exe`
- **Project file**: `MosaicTools.csproj` (NOT MosaicToolsCSharp.csproj!)
- **Working directory**: `c:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp`
- **Output**: `bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe`

### Full Build & Run Command
```bash
EXE="C:/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp/bin/Release/net8.0-windows10.0.19041.0/win-x64/publish/MosaicTools.exe" && taskkill //IM MosaicTools.exe //F 2>/dev/null; for i in $(seq 1 20); do [ ! -f "$EXE" ] && break; rm -f "$EXE" 2>/dev/null && break; sleep 0.5; done && cd /c/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp && /c/Users/erik.richter/Desktop/dotnet/dotnet.exe publish -c Release -r win-x64 --self-contained 2>&1 && start "" "$EXE"
```

### Quick Debug Build (syntax check only)
```bash
cd /c/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp && /c/Users/erik.richter/Desktop/dotnet/dotnet.exe build MosaicTools.csproj
```

### Shell Gotchas
- Use `//IM` and `//F` (double slash) for taskkill — MSYS2/Git Bash converts single `/I` to a file path
- PowerShell commands need `-ExecutionPolicy Bypass` — default policy blocks module loading

---

## GitHub & Releases

- **Repo:** https://github.com/erichter2018/MosaicTools (public, required for auto-update)
- **gh CLI:** `"C:\Users\erik.richter\Desktop\GH CLI\gh.exe"`

### Auto-Update System

Updates via GitHub Releases using a rename trick (no installer, no admin). Uses ZIP files to avoid corporate antivirus blocking exe downloads. Both .zip and .exe are published for backwards compatibility (pre-2.5.1 clients only look for .exe).

### Publishing a Release

When user says "create a release" or "publish release":

1. **Update version** in `MosaicToolsCSharp/MosaicTools.csproj` — ALL THREE fields must match:
   ```xml
   <Version>X.Y.Z</Version>
   <AssemblyVersion>X.Y.Z.0</AssemblyVersion>
   <FileVersion>X.Y.Z.0</FileVersion>
   ```
   `AssemblyVersion` is used for update comparison — mismatch causes update loops!

2. **Update WhatsNew.txt** — prepend new version section at top of `MosaicToolsCSharp/WhatsNew.txt`:
   - Run `git log v{previous}..HEAD --oneline` to see changes
   - Keep ALL entries back to the last x.0.0 release (users who skip updates see all missed changes)

3. **Commit and push:**
   ```bash
   git add -A && git commit -m "vX.Y.Z: Release notes here" && git push
   ```

4. **Kill + build** (do NOT start the app — locked exe prevents zipping):
   ```bash
   EXE="C:/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp/bin/Release/net8.0-windows10.0.19041.0/win-x64/publish/MosaicTools.exe" && taskkill //IM MosaicTools.exe //F 2>/dev/null; for i in $(seq 1 20); do [ ! -f "$EXE" ] && break; rm -f "$EXE" 2>/dev/null && break; sleep 0.5; done && cd /c/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp && /c/Users/erik.richter/Desktop/dotnet/dotnet.exe publish -c Release -r win-x64 --self-contained 2>&1
   ```

5. **Create ZIP** (MUST use `-ExecutionPolicy Bypass`):
   ```bash
   powershell -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe' -DestinationPath 'C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.zip' -Force"
   ```

6. **Create GitHub release** with both zip and exe:
   ```bash
   "C:\Users\erik.richter\Desktop\GH CLI\gh.exe" release create vX.Y.Z "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.zip" "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe" --title "vX.Y.Z" --notes "Release notes here"
   ```

---

## Architecture Overview

### Project Structure
```
MosaicToolsCSharp/
├── Program.cs                  # Entry point, mutex, exe normalization
├── App.cs                      # Global state (IsHeadless flag)
├── MosaicTools.csproj          # Project file with version info
├── WhatsNew.txt                # Embedded changelog for What's New popup
├── Services/
│   ├── Configuration.cs        # Settings (JSON persistence to %LOCALAPPDATA%\MosaicTools\)
│   ├── AutomationService.cs    # FlaUI-based UI Automation for Mosaic/Clario
│   ├── IMosaicReader.cs        # Interface: read-only Mosaic state
│   ├── IMosaicCommander.cs     # Interface: Mosaic UI commands
│   ├── HidService.cs           # PowerMic/SpeechMike USB HID communication
│   ├── KeyboardService.cs      # Global hotkey registration
│   ├── NativeWindows.cs        # Win32 API (window mgmt, keyboard sim, focus)
│   ├── UpdateService.cs        # GitHub release auto-update
│   ├── NoteFormatter.cs        # Clario note parsing → critical findings text
│   ├── GetPriorService.cs      # Prior study text formatting
│   ├── OcrService.cs           # Windows.Media.Ocr for series/image capture
│   ├── ClipboardService.cs     # Clipboard operations (STA-safe)
│   ├── AudioService.cs         # Beep sounds for dictation feedback
│   ├── Logger.cs               # File logging to mosaic_tools_log.txt
│   ├── InputBox.cs             # Simple text input dialog
│   ├── AidocService.cs         # Aidoc widget scraping (positive finding detection)
│   ├── AidocFindingVerifier.cs # Study-type relevance filtering for Aidoc findings
│   ├── SttService.cs           # Custom STT orchestrator (audio capture → provider → paste)
│   ├── ISttProvider.cs         # STT provider interface
│   ├── DeepgramProvider.cs     # Deepgram Nova-3 / Nova-3 Medical WebSocket STT
│   ├── AssemblyAIProvider.cs   # AssemblyAI streaming STT
│   ├── CortiProvider.cs        # Corti Solo medical STT (Opus/OAuth)
│   ├── StreamingOggOpusWriter.cs # Opus audio encoding for Corti
│   ├── WebmOpusMuxer.cs        # WebM muxer for Corti audio stream
│   ├── RadAiService.cs         # [RadAI] REST API for AI impressions (tagged for removal)
│   ├── RecoMdService.cs        # RecoMD best-practice recommendations
│   ├── CorrelationService.cs   # Rainbow mode: findings↔impression correlation
│   ├── TemplateDatabase.cs     # SQLite template DB for report change detection
│   ├── RvuCounterService.cs    # RVUCounter named-pipe integration
│   ├── PipeService.cs          # Named pipe server for external tool communication
│   ├── ConnectivityService.cs  # Network connectivity monitoring
│   └── CriticalStudyEntry.cs   # Data model for critical study tracker
└── UI/
    ├── ActionController.cs     # Central coordinator for all actions (3500+ lines)
    ├── MainForm.cs             # Main widget bar + toast system + RVU display
    ├── SettingsFormNew.cs      # Settings dialog (sidebar nav, search, 11 sections)
    ├── Settings/               # Settings sections (one per tab)
    │   ├── SettingsSection.cs  # Base class for all sections
    │   ├── ProfileSection.cs
    │   ├── DesktopSection.cs
    │   ├── KeysButtonsSection.cs
    │   ├── TextTemplatesSection.cs
    │   ├── AlertsSection.cs
    │   ├── ReportDisplaySection.cs
    │   ├── BehaviorSection.cs
    │   ├── RvuMetricsSection.cs
    │   ├── SttSection.cs
    │   ├── ExperimentalSection.cs
    │   └── ReferenceSection.cs
    ├── FloatingToolbarForm.cs  # Configurable IV button grid
    ├── IndicatorForm.cs        # Recording state indicator dot
    ├── ClinicalHistoryForm.cs  # Notification box (clinical history + alerts)
    ├── ImpressionForm.cs       # Auto-show impression after Process Report
    ├── ReportPopupForm.cs      # Report viewer (changes/rainbow/orphan modes)
    ├── TranscriptionForm.cs    # Live STT transcription overlay
    ├── WhatsNewForm.cs         # Post-update changelog popup
    ├── CriticalStudiesPopup.cs # Critical studies tracker list
    ├── RvuPopupForm.cs         # RVU metrics hover popup
    ├── RadAiOverlayForm.cs     # [RadAI] Impression display popup
    ├── PickListPopupForm.cs    # Pick list selection popup
    ├── KeyMappingsDialog.cs    # Hotkey/mic button mapping editor
    ├── ButtonStudioDialog.cs   # FloatingToolbar button editor
    ├── MacroEditorForm.cs      # Macro definition editor
    ├── PickListEditorForm.cs   # Pick list definition editor
    ├── AudioSetupForm.cs       # Mic gain calibrator
    ├── ConnectivityDetailsForm.cs # Network status details
    ├── ScreenHelper.cs         # Multi-monitor bounds/offscreen detection
    └── LayeredWindowHelper.cs  # Win32 layered window for transparent overlays
```

### Key Data Flow

**Startup:** `Program.Main()` → mutex check → `Configuration.Load()` → `MainForm` → `ActionController` → starts services, checks updates

**Action Triggering (3 sources):**
1. PowerMic/SpeechMike buttons → `HidService.ButtonPressed` → `ActionController.TriggerAction()`
2. Keyboard hotkeys → `KeyboardService` → `ActionController.TriggerAction()`
3. Windows messages → `MainForm.WndProc()` → `ActionController.TriggerAction()`

**Action Execution:** All actions queued to dedicated STA thread → `ActionController.ExecuteAction()` → `Perform*()` methods. Focus saved/restored around Mosaic interactions.

### Available Actions (defined in Configuration.Actions)
| Action | Description |
|--------|-------------|
| `Get Prior` | Extract prior study from InteleViewer, format, paste to Mosaic |
| `Critical Findings` | Scrape Clario exam note, format, paste to Mosaic |
| `Capture Series/Image` | OCR screen for series/image numbers, paste to Mosaic |
| `Process Report` | Alt+P in Mosaic, optional auto-stop dictation, smart scroll |
| `Sign Report` | Alt+F in Mosaic |
| `Start/Stop Recording` | Alt+R in Mosaic, with beep feedback |
| `System Beep` | Toggle dictation state tracking with audio feedback |
| `Show Report` | Alt+C to copy report, show in popup |
| `Create Impression` | Generate impression section |
| `Discard Study` | Discard current study |
| `Show Pick Lists` | Show pick list selection popup |
| `Cycle Window/Level` | Send window/level keys to InteleViewer |
| `Create Critical Note` | Create Clario communication note for stroke cases |
| `RadAI Impression` | [RadAI] Generate and show AI impression |
| `Trigger RecoMD` | Send report to RecoMD for recommendations |
| `Paste RecoMD` | Paste RecoMD recommendations into report |

### Configuration System
- **Path:** `%LOCALAPPDATA%\MosaicTools\MosaicToolsSettings.json`
- **Serialization:** `System.Text.Json` with `[JsonPropertyName]` — all properties auto-serialize
- **Settings UI:** `SettingsFormNew` with 11 section classes, each with `LoadSettings()`/`SaveSettings()`

### Background Timers
1. **Sync Timer (250ms)** — registry-based dictation state polling for indicator
2. **Scrape Timer (configurable, default 3s)** — Mosaic UI scraping via FlaUI
   - Tracks accession changes, drafted state, clinical history, patient info
   - Speeds up to 1s when searching for impression after Process Report

### External Dependencies
- **FlaUI** (FlaUI.UIA3) — UI Automation wrapper
- **HidSharp** — USB HID for PowerMic/SpeechMike
- **NAudio** — Audio capture for custom STT
- **Concentus** — Opus audio encoding for Corti STT provider
- **Windows.Media.Ocr** — Built-in Windows OCR API

### Headless Mode
Launch with `-headless`: no widget bar, no hotkeys, system tray icon for settings. PowerMic, floating toolbar, and toasts still work.

### Open in Clario (XML File Drop)
Double-click in Critical Studies popup writes XML to `C:\MModal\FluencyForImaging\Reporting\XML\IN` — Clario watches this folder and opens the study.

### Common Debugging
- **Log file:** `mosaic_tools_log.txt` in exe directory
- **Debug scrape:** Hold Win key + trigger Critical Findings → shows raw data dialog
- **Toast messages:** Bottom-right stacking notifications

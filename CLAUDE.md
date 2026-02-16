# MosaicTools Project

## For New Claude Sessions

This file is automatically loaded at the start of every Claude Code session. It contains everything needed to understand this codebase without re-exploring. If you're a new Claude session:

1. **Don't re-analyze** - The architecture section below has the full codebase structure
2. **Check recent commits** - Run `git log --oneline -10` to see what changed recently
3. **Read the relevant service** - Jump straight to the file you need (paths are in Architecture section)
4. **Build command is below** - Just use it, don't ask for confirmation

If making significant changes to the codebase structure, update the Architecture section at the bottom of this file so future sessions stay current.

---

## ClarioIgnore - Separate Project

**IMPORTANT:** ClarioIgnore is a separate tool in the `ClarioIgnore/` folder. It is NOT part of MosaicTools releases. When publishing MosaicTools releases, do NOT include ClarioIgnore. ClarioIgnore code can be committed to the repo but should never be published as a GitHub release.

---

## Build Instructions

**IMPORTANT: When user says "build" or "rebuild", OR when you finish making code changes, just run the full build command immediately without asking for confirmation. This includes taskkill, compile, and starting the app - do it all automatically.**

### CRITICAL: dotnet SDK Location
The .NET SDK is **NOT in system PATH**. It's located on the Desktop:
```
c:\Users\erik.richter\Desktop\dotnet\dotnet.exe
```

### Project Details
- **Project file**: `MosaicTools.csproj` (NOT MosaicToolsCSharp.csproj!)
- **Working directory**: `c:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp`
- **Output location**: `bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe`

### Full Build & Run Command (Use This!)
```bash
EXE="C:/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp/bin/Release/net8.0-windows10.0.19041.0/win-x64/publish/MosaicTools.exe" && taskkill //IM MosaicTools.exe //F 2>/dev/null; for i in $(seq 1 20); do [ ! -f "$EXE" ] && break; rm -f "$EXE" 2>/dev/null && break; sleep 0.5; done && cd /c/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp && /c/Users/erik.richter/Desktop/dotnet/dotnet.exe publish -c Release -r win-x64 --self-contained 2>&1 && start "" "$EXE"
```

**IMPORTANT:** Use `//IM` and `//F` (double slash) — MSYS2/Git Bash converts single `/I` to a file path, silently breaking taskkill.

### Quick Debug Build (syntax checking only)
```bash
cd /c/Users/erik.richter/Desktop/MosaicTools/MosaicToolsCSharp && /c/Users/erik.richter/Desktop/dotnet/dotnet.exe build MosaicTools.csproj
```

### Common Issues
1. **"Project file does not exist"** - Use `MosaicTools.csproj`, not `MosaicToolsCSharp.csproj`
2. **"dotnet is not recognized"** - Use full path as shown above

---

## GitHub Repository

- **URL:** https://github.com/erichter2018/MosaicTools
- **Account:** erichter@gmail.com
- **Visibility:** Public (required for auto-update to work)

### gh CLI Location
```
"C:\Users\erik.richter\Desktop\GH CLI\gh.exe"
```

---

## Auto-Update System

The app auto-updates via GitHub Releases using a rename trick (no batch files, no installer, no admin required).

**IMPORTANT: Releases use ZIP files** to avoid corporate security/antivirus blocking direct exe downloads.

### How it works
1. On startup, checks `https://api.github.com/repos/erichter2018/MosaicTools/releases/latest`
2. If newer version found, downloads `MosaicTools.zip` to temp location
3. Extracts `MosaicTools.exe` from the zip to `MosaicTools_new.exe`
4. Renames running exe to `MosaicTools_old.exe` (Windows allows renaming running files)
5. Renames new exe to `MosaicTools.exe`
6. Shows toast with "Restart Now" button
7. On next startup, deletes `_old.exe` and any leftover zip files

### Publishing a new release

When user says "create a release" or "publish release vX.X":

1. **Update version** in `MosaicToolsCSharp/MosaicTools.csproj` (ALL THREE fields!):
   ```xml
   <Version>2.5.1</Version>
   <AssemblyVersion>2.5.1.0</AssemblyVersion>
   <FileVersion>2.5.1.0</FileVersion>
   ```

2. **Update WhatsNew.txt** - Prepend new version section:
   - Run: `git log v{previous}..HEAD --oneline` to see changes since last release
   - Summarize commits into brief bullet points (1 line per feature)
   - Group minor fixes as "Bug fixes"
   - Add version header and bullets to top of `MosaicToolsCSharp/WhatsNew.txt`
   - Example format:
     ```
     2.5.5
     - What's New popup shows new features after updates
     - Bug fixes
     ```

3. **Commit and push** the version change:
   ```bash
   git add -A && git commit -m "v2.5.1: Release notes here" && git push
   ```

4. **Build** the release exe:
   ```powershell
   cd C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp
   c:\Users\erik.richter\Desktop\dotnet\dotnet.exe publish -c Release -r win-x64 --self-contained
   ```

5. **Create the ZIP file** containing MosaicTools.exe:
   ```powershell
   Compress-Archive -Path "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe" -DestinationPath "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.zip" -Force
   ```

6. **Create the GitHub release** with BOTH zip and exe (for backwards compatibility):
   ```bash
   "C:\Users\erik.richter\Desktop\GH CLI\gh.exe" release create v2.5.1 "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.zip" "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe" --title "v2.5.1" --notes "Release notes here"
   ```

**Why both files?** Old versions (pre-2.5.1) only look for .exe, new versions prefer .zip but fall back to .exe.

### Direct Download Links
```
https://github.com/erichter2018/MosaicTools/releases/latest/download/MosaicTools.zip
https://github.com/erichter2018/MosaicTools/releases/latest/download/MosaicTools.exe
```

### Settings
- **Auto-update** checkbox in General tab (ON by default)
- **Check for Updates** button for manual checks
- Update check has 15 second timeout to prevent hanging on blocked networks

### IMPORTANT: Version Numbers
When releasing, you MUST update ALL THREE version fields in the csproj:
```xml
<Version>2.5.1</Version>
<AssemblyVersion>2.5.1.0</AssemblyVersion>
<FileVersion>2.5.1.0</FileVersion>
```
The auto-update uses `AssemblyVersion` for comparison - if it doesn't match the release tag, users get stuck in an update loop!

---

## Architecture Overview

### Project Structure
```
MosaicToolsCSharp/
├── Program.cs              # Entry point, mutex, exe normalization
├── App.cs                  # Global state (IsHeadless flag)
├── MosaicTools.csproj      # Project file with version info
├── Services/               # Business logic layer
│   ├── Configuration.cs    # Settings (JSON persistence to AppData)
│   ├── ActionController.cs # Central coordinator for all actions
│   ├── AutomationService.cs# UI Automation (FlaUI) for Mosaic/Clario
│   ├── HidService.cs       # PowerMic USB HID communication
│   ├── KeyboardService.cs  # Global hotkey registration
│   ├── NativeWindows.cs    # Win32 API (window mgmt, keyboard sim)
│   ├── UpdateService.cs    # GitHub release auto-update
│   ├── NoteFormatter.cs    # Clario note parsing → critical findings
│   ├── GetPriorService.cs  # Prior study text formatting
│   ├── OcrService.cs       # Windows.Media.Ocr for series/image capture
│   ├── ClipboardService.cs # Clipboard operations (STA-safe)
│   ├── AudioService.cs     # Beep sounds for dictation feedback
│   ├── Logger.cs           # File logging to mosaic_tools_log.txt
│   └── InputBox.cs         # Simple text input dialog
├── WhatsNew.txt            # Embedded changelog for What's New popup
└── UI/                     # WinForms presentation layer
    ├── MainForm.cs         # Main widget bar + toast system
    ├── SettingsForm.cs     # Configuration dialog (tabbed)
    ├── FloatingToolbarForm.cs  # Configurable button grid
    ├── IndicatorForm.cs    # Recording state indicator light
    ├── ClinicalHistoryForm.cs  # Clinical history display window
    ├── ImpressionForm.cs   # Auto-show impression during drafting
    ├── ReportPopupForm.cs  # Full report viewer popup
    └── WhatsNewForm.cs     # Post-update changelog popup
```

### Key Data Flow

**Startup Flow:**
```
Program.Main()
  → Mutex check (single instance)
  → NormalizeExecutableName() (ensures MosaicTools.exe name)
  → Configuration.Load() (from %LOCALAPPDATA%\MosaicTools\)
  → MainForm created
    → ActionController created (coordinates everything)
    → OnFormShown: starts services, checks updates
```

**Action Triggering (3 input sources):**
```
1. PowerMic buttons → HidService.ButtonPressed → ActionController.TriggerAction()
2. Keyboard hotkeys → KeyboardService → ActionController.TriggerAction()
3. Windows messages → MainForm.WndProc() → ActionController.TriggerAction()
```

**Action Execution:**
- All actions queued to dedicated STA thread (required for clipboard/SendKeys)
- `ActionController.ExecuteAction()` dispatches to specific `Perform*()` methods
- Focus saved/restored around Mosaic interactions

### Available Actions (defined in Configuration.cs)
| Action | Description |
|--------|-------------|
| `Get Prior` | Extract prior study from InteleViewer, format, paste to Mosaic |
| `Critical Findings` | Scrape Clario for exam note, format, paste to Mosaic |
| `Capture Series` | OCR screen for series/image numbers, paste to Mosaic |
| `Process Report` | Alt+P in Mosaic, optional auto-stop dictation, smart scroll |
| `Sign Report` | Alt+F in Mosaic |
| `Toggle Record` | Alt+R in Mosaic, with beep feedback |
| `System Beep` | Toggle dictation state tracking with audio feedback |
| `Show Report` | Alt+C to copy report, show in popup |

### Key Services Detail

**AutomationService (FlaUI-based):**
- `FindClarioWindow()` - Locates Chrome with "Clario - Worklist"
- `GetExamNoteElements()` - Searches DataItem elements for "EXAM NOTE"
- `GetFinalReportFast()` - Fast scrape of Mosaic's ProseMirror editor
- Tracks: `LastFinalReport`, `LastAccession`, `LastDraftedState`, `LastDescription`, `LastPatientGender`

**HidService (HidSharp library):**
- Connects to Nuance PowerMic (Vendor IDs: 0x0554, 0x0558)
- Button events: `ButtonPressed`, `RecordButtonStateChanged` (for PTT mode)
- Runs on background thread with non-blocking reads

**NativeWindows (Win32 interop):**
- Window activation: `ActivateMosaicForcefully()` - aggressive multi-attempt activation
- Keyboard simulation: `SendAltKey()`, `SendHotkey()`, `KeyUpModifiers()`
- Dictation state: `IsMicrophoneActiveFromRegistry()` - reads Windows mic consent store
- Focus management: `SavePreviousFocus()`, `RestorePreviousFocus()`

**NoteFormatter:**
- Parses Clario exam notes (e.g., "Transferred Smith to Jones at 3:45 PM...")
- Extracts: contact name, timestamp, timezone
- Outputs: "Critical findings were discussed with and acknowledged by {name} at {time} on {date}."
- Filters out the user's own name via `Configuration.DoctorName`

### Configuration System
- **Path:** `%LOCALAPPDATA%\MosaicTools\MosaicToolsSettings.json`
- **Migration:** Auto-migrates from old exe-relative location
- **First run:** Shows onboarding dialog for doctor name
- **Key settings:**
  - `DoctorName` - Used to filter names in note parsing
  - `ActionMappings` - Maps actions to hotkeys and mic buttons
  - `FloatingButtons` - Configurable button grid definition
  - `ScrapeMosaicEnabled` - Background polling of Mosaic state
  - Feature flags: `ShowClinicalHistory`, `ShowImpression`, `GenderCheckEnabled`, etc.

### Background Timers
1. **Sync Timer (250ms)** - Registry-based dictation state polling for indicator
2. **Scrape Timer (configurable, default 3s)** - Mosaic UI scraping when enabled
   - Tracks accession changes, drafted state, clinical history
   - Speeds up to 1s when searching for impression after Process Report

### UI Windows
All forms are borderless, topmost, draggable:
- **MainForm** - 160x40px "Mosaic Tools" bar, click for settings
- **FloatingToolbarForm** - Dynamic button grid from `FloatingButtons` config
- **IndicatorForm** - Small red/gray dot showing recording state
- **ClinicalHistoryForm** - Shows extracted clinical history, color-coded warnings
- **ImpressionForm** - Auto-shows impression section when report drafted

### Headless Mode
Launch with `-headless` flag:
- No widget bar (invisible)
- No hotkeys registered
- System tray icon for settings access
- PowerMic, floating toolbar, and toasts still work

### External Dependencies
- **FlaUI** - UI Automation wrapper (NuGet: FlaUI.UIA3)
- **HidSharp** - USB HID communication (NuGet: HidSharp)
- **Windows.Media.Ocr** - Built-in Windows OCR API

### Open in Clario (XML File Drop)
The "Open in Clario" feature (double-click in Critical Studies popup) works by writing an XML file to a Fluency watch folder. Clario monitors this folder and opens the study.

- **Folder:** `C:\MModal\FluencyForImaging\Reporting\XML\IN`
- **Method:** `AutomationService.OpenStudyInClario(accession, mrn)` (line ~554)
- **File format:** `openreport{unixTimestamp}.{pid}.xml`
- **XML content:**
  ```xml
  <Message>
    <Type>OpenReport</Type>
    <AccessionNumbers>
      <AccessionNumber>{accession}</AccessionNumber>
    </AccessionNumbers>
    <MedicalRecordNumber>{mrn}</MedicalRecordNumber>
  </Message>
  ```
- **Requires:** Both accession and MRN (scraped from Mosaic UI by `AutomationService`)
- **Availability check:** `IsXmlFolderAvailable()` checks if the IN folder exists
- **Called from:** `CriticalStudiesPopup.ListBox_DoubleClick`

### Common Debugging
- **Log file:** `mosaic_tools_log.txt` in exe directory
- **Debug scrape:** Hold Win key + trigger Critical Findings → shows raw data dialog
- **Toast messages:** Bottom-right stacking notifications

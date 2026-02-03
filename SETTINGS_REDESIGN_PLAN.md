# Settings UI Redesign Plan

## Goal
Replace tabbed settings with VS Code-style searchable interface. Move Keys and IV Buttons to separate dialogs.

## New Layout

```
┌─────────────────────────────────────────────────────┐
│  Mosaic Tools Settings v3.0.X                       │
│                                                      │
│  🔍 [Search settings...]                            │
│ ─────────────────────────────────────────────────── │
│  ┌─ Profile ─────────────────────────────────────┐ │
│  │ Doctor Name: [              ]                  │ │
│  │ InteleViewer Hotkey: [    ]                    │ │
│  └────────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ Display ──────────────────────────────────────┐ │
│  │ ☑ Show Floating Toolbar                        │ │
│  │ ☑ Show Recording Indicator                     │ │
│  │ ☑ Hide Indicator When No Study                 │ │
│  │ ☑ Show Tooltips                                │ │
│  │ Report Font: [Consolas] Size: [10]             │ │
│  │ ☑ Transparent Report Window  [▬▬●───] 75%     │ │
│  │ ☑ Highlight Report Changes  [Color] [Alpha]   │ │
│  └────────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ RVU Counter ───────────────────────────────── │
│  │ ☑ Enable RVU Counter                          │ │
│  │ Show Metrics: ☑Total ☑Per Hour ☐Current Hour │ │
│  │ Overflow Layout: [Horizontal ▼]               │ │
│  │ ☑ Enable Goal  Goal/Hour: [10.0]              │ │
│  │ Counter Path: [C:\...\RVUCounter.exe]         │ │
│  └───────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ Dictation ────────────────────────────────────┐ │
│  │ ☑ Start Beep [▬▬●───] 50%                      │ │
│  │ ☑ Stop Beep  [▬●─────] 25%                     │ │
│  │ Pause Duration: [3] seconds                    │ │
│  │ ☑ Auto-stop after Process Report               │ │
│  │ ☑ Dead Man Switch                              │ │
│  └────────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ Text Automation ──────────────────────────────┐ │
│  │ Critical Template: [...] Series: [...] [...]   │ │
│  │ ☑ Separate Pasted Items with Blank Lines       │ │
│  │ ☑ Enable Macros  ☑ Strip Blank Lines           │ │
│  │ ☑ Enable Pick Lists  ☑ Skip Single Match       │ │
│  └────────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ Mosaic Integration ───────────────────────────┐ │
│  │ ☑ Scrape Mosaic  Interval: [3] seconds         │ │
│  │ ☑ Restore Focus After Actions                   │ │
│  │ ☑ Scroll to Bottom After Process                │ │
│  │ Scroll Thresholds: [10] [20] [30] lines        │ │
│  │ ☐ Ignore Inpatient Drafted Studies              │ │
│  └────────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ Alerts & Notifications ───────────────────────┐ │
│  │ ☑ Show Clinical History  ☑ Always Show         │ │
│  │ ☑ Hide When No Study  ☑ Auto-fix Format        │ │
│  │ ☑ Show Drafted Indicator                        │ │
│  │ ☑ Show Template Mismatch Warnings               │ │
│  │ ☑ Gender Check Enabled                          │ │
│  │ ☑ Stroke Detection  ☑ Use Clinical History     │ │
│  │   ☑ Click to Create Note  ☐ Auto Create        │ │
│  │ ☑ Track Critical Studies                        │ │
│  │ ☑ Show Impression Window                        │ │
│  │ ☑ Show Report After Process                     │ │
│  │ ☑ Show Line Count Toast                         │ │
│  └────────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ Network & Updates ────────────────────────────┐ │
│  │ ☑ Auto-Update from GitHub                       │ │
│  │ ☑ Monitor Network Connectivity                  │ │
│  │   Interval: [30] sec  Timeout: [5] sec          │ │
│  └────────────────────────────────────────────────┘ │
│                                                      │
│  [Configure Hotkeys...] [Configure IV Buttons...]   │
│ ─────────────────────────────────────────────────── │
│  [Help]       v3.0.6       [Save]  [Cancel]         │
└─────────────────────────────────────────────────────┘
```

## Settings Organization

### Profile
- Doctor Name
- InteleViewer Hotkey

### Display
- Show Floating Toolbar
- Show Recording Indicator
- Hide Indicator When No Study
- Show Tooltips
- Report Font (Family, Size)
- Transparent Report Window (Enabled, Opacity %)
- Highlight Report Changes (Enabled, Color, Alpha)

### RVU Counter
- Enable RVU Counter
- Display Metrics (Total, Per Hour, Current Hour, Prior Hour, Estimated Total)
- Overflow Layout (Horizontal, Vertical Stack, Hover Popup, Carousel)
- Enable Goal
- Goal Per Hour
- Counter Path

### Dictation
- Start Beep (Enabled, Volume)
- Stop Beep (Enabled, Volume)
- Pause Duration
- Auto-stop After Process
- Dead Man Switch

### Text Automation
- Critical Findings Template
- Series Template
- Comparison Template
- Separate Pasted Items
- Macros (Enabled, Strip Blank Lines)
- Pick Lists (Enabled, Skip Single Match, Keep Open)

### Mosaic Integration
- Scrape Mosaic (Enabled, Interval)
- Restore Focus After Actions
- Scroll to Bottom After Process
- Scroll Thresholds (1, 2, 3)
- Ignore Inpatient Drafted Studies (Enabled, Chest Only vs All XR)

### Alerts & Notifications
- Clinical History (Show, Always Show, Hide When No Study, Auto-fix)
- Show Drafted Indicator
- Show Template Mismatch
- Gender Check
- Stroke Detection (Enabled, Use Clinical History, Click to Create, Auto Create)
- Track Critical Studies
- Show Impression
- Show Report After Process
- Line Count Toast

### Network & Updates
- Auto-Update
- Connectivity Monitor (Enabled, Interval, Timeout)

## Separate Dialogs

### Keys Configuration Dialog (KeysConfigDialog.cs)
- Moved from "Keys" tab
- Hotkey mapping table
- Action assignments

### IV Buttons Configuration Dialog (IVButtonsConfigDialog.cs)
- Moved from "IV Buttons" tab
- Button grid editor
- Icon/label/keystroke/action configuration

## Search Implementation

The search box filters settings in real-time:
1. User types in search box
2. For each section and control:
   - Check if search term matches:
     - Control label text
     - Tooltip text
     - Section name
3. Hide sections/controls that don't match
4. Show only matching items
5. Highlight search term in results (optional)

Search is case-insensitive and matches partial strings.

## Implementation Tasks

1. ✓ Analyze current settings structure
2. ✓ Design new settings organization (this file)
3. ✓ Create KeysConfigDialog.cs (extract from Keys tab)
4. ✓ Create IVButtonsConfigDialog.cs (extract from IV Buttons tab)
5. ✓ Create SettingsFormNew.cs:
   - Removed TabControl
   - Added search TextBox at top
   - Created scrollable Panel with GroupBox sections
   - Implemented search filtering logic
   - Added buttons to open Keys and IV Buttons dialogs
6. ✓ Test and build - Application running successfully!

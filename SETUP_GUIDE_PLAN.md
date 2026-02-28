# Setup Guide / Onboarding Wizard

## Context

Users don't discover MosaicTools features and don't look through all 11 settings sections. Current onboarding is just a name prompt. We need a proper feature catalog with short demos and inline toggles so users can configure their workspace during first launch and revisit anytime.

## Architecture

- **WebView2 + embedded HTML** (proven pattern from AudioSetupForm/MicCalibrator)
- **900x700 dark-themed modal** — feature catalog (scrollable, all cards visible)
- **Batch save** on "Done" button — nothing changes until user commits
- **Auto-detects** PowerMic/SpeechMike and Aidoc
- **Placeholder preview areas** in each card (GIFs/animations added later)
- **Pro tip callouts** with workflow recommendations

## Files

| File | Action | Est. Lines |
|------|--------|------------|
| `UI/SetupGuideForm.cs` | NEW | ~250 |
| `SetupGuide.html` | NEW (embedded resource) | ~1000 |
| `Services/Configuration.cs` | MODIFY | ~15 lines |
| `UI/MainForm.cs` | MODIFY | ~40 lines |
| `MosaicTools.csproj` | MODIFY | 1 line |

## Feature Cards (5 categories, ~14 cards)

### 1. Getting Started
- **Your Profile** — Name input, timezone selector
- **Microphone Device** ★ MUST-KNOW — Auto-detect PowerMic/SpeechMike, show status badge, device preference dropdown

### 2. Clinical Alerts
- **Clinical History Box** ★ MUST-KNOW — Main toggle + sub-toggles (age/gender/drafted/auto-fix). Pro tip: "Enable everything — the box stays out of the way and saves you from missing critical info."
- **Stroke Detection** — Toggle. Pro tip: "Enable for purple border alerts on stroke cases."
- **Aidoc Integration** — Toggle, auto-detect if running. Pro tip: "If your site runs Aidoc, enable this. Disable pulse animation in Aidoc settings first."

### 3. Report Tools
- **Report Display** — Toggle changes + rainbow. Pro tip: "Turn on Rainbow mode for best findings/impression correlation."
- **Impression Fixer** — Toggle. Pro tip TBD.
- **RecoMD Recommendations** — Toggle. Pro tip TBD.

### 4. Dictation
- **Recording Beeps** — Toggle start/stop beeps with volume sliders
- **Custom STT** — Toggle, provider dropdown, API key. Pro tip: "Only if you're frustrated with Mosaic's dictation. Requires API key + internet."

### 5. Advanced
- **Pick Lists** — Toggle
- **Macros** — Toggle
- **RVU Counter** — Toggle
- **Floating Toolbar** — Toggle

**Excluded:** Connectivity Monitor

## Implementation Steps

### Step 1: Skeleton (SetupGuideForm + minimal HTML)
1. Create `SetupGuideForm.cs` following `AudioSetupForm.cs` pattern:
   - WebView2 hosting, DWM dark title bar, user data folder at `%LOCALAPPDATA%\MosaicTools\WebView2\`
   - `BuildInitData()` — serializes current config + detection results as JSON
   - `OnWebMessageReceived()` — handles `save` and `cancel` messages from JS
   - `ApplySettings(JsonElement)` — maps JS camelCase properties to C# config properties
   - Fallback to InputBox if WebView2 unavailable (first-run only)
2. Create minimal `SetupGuide.html` with Getting Started category only
3. Add `<EmbeddedResource Include="SetupGuide.html" />` to csproj
4. **Build & test**: verify HTML loads, C#→JS init data arrives

### Step 2: First-Run Integration
1. Add `HasCompletedSetup` (bool, default false) to `Configuration.cs`
2. Replace `ShowOnboarding()` — remove MessageBox/InputBox, just create default config silently
3. In `MainForm.OnFormShown()` AFTER `_controller.Start()`: if `!config.HasCompletedSetup`, call `ShowSetupGuide(isFirstRun: true)`
4. Add `ShowSetupGuide()` method with post-save refresh (same pattern as `SettingsFormNew.SaveAndClose()`)
5. Add "Setup Guide" to both widget bar and tray icon context menus
6. **Build & test**: delete settings file → wizard appears; menu item works for re-access

### Step 3: Full Content
1. Build all 5 categories and ~14 cards in HTML
2. Wire all toggles, inputs, selects in JS `populateFields()` and `collectSettings()`
3. Complete `ApplySettings()` C# mapping for all ~30 properties
4. **Build & test**: toggle features, Done, verify settings persist

### Step 4: Auto-Detection
1. Call `HidService.GetAvailableDevices()` in `BuildInitData()`
2. Check for Aidoc window existence
3. JS shows green "Detected!" badges, auto-suggests toggle state
4. **Build & test**: with/without PowerMic, with/without Aidoc

### Step 5: Polish
1. CSS transitions, hover effects, toggle switch animations
2. "Enable All" shortcut in Clinical History pro tip
3. Conditional sub-control visibility (STT provider→API key field, Clinical History→sub-toggles)
4. Loading spinner until `initWizard()` called
5. Font fallback (Segoe UI if Google Fonts unavailable offline)

## Key Patterns to Follow

- **AudioSetupForm.cs** — WebView2 hosting, virtual host mapping, dark title bar, resource loading
- **MicCalibrator.html** — CSS design tokens (--bg, --surface, --border, --text, etc.), card styling
- **SettingsFormNew.SaveAndClose()** — Post-save refresh call sequence (RefreshServices, ToggleFloatingToolbar, etc.)
- **HidService.GetAvailableDevices()** — Static method for device detection

## WebView2 <-> C# Communication

**C# -> JS** (on navigation complete):
```js
window.initWizard({ isFirstRun, detectedDevices, aidocDetected, currentSettings })
```

**JS -> C#** (on Done):
```js
window.chrome.webview.postMessage({ type: "save", settings: { ... } })
```

## Edge Cases
- WebView2 not installed -> fallback to InputBox name prompt (first run only)
- User closes via X -> same as Cancel, `HasCompletedSetup` stays false, wizard re-shows next launch
- Empty name -> defaults to "Radiologist"
- No internet -> Google Fonts fallback to Segoe UI, all content still works (embedded)
- `ShowClinicalHistory` toggle ON -> also sets `AlwaysShowClinicalHistory` = true

## Verification
1. Delete `%LOCALAPPDATA%\MosaicTools\MosaicToolsSettings.json` -> launch -> wizard appears
2. Fill in name, toggle features, click Done -> settings file created with correct values
3. Launch again -> wizard does NOT appear (HasCompletedSetup = true)
4. Right-click widget bar -> "Setup Guide" -> wizard opens with current settings populated
5. Toggle features in wizard, Done -> features take effect (clinical history window appears, etc.)
6. Test with PowerMic connected -> "Detected!" badge shows

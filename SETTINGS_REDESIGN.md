# Settings Form Redesign - Windows 11 Style

## Design Goals
1. **Side navigation** - Left sidebar with section list, content scrolls on the right
2. **Search** - Filters entire sections (not individual items)
3. **Smooth scrolling** - Animated scroll to sections, smooth mouse wheel
4. **Modern appearance** - Dark Fluent-inspired, clean spacing, no clutter
5. **Tooltips on labels** - Hover any label for help (no question marks)

---

## Visual Layout

```
┌─────────────────────────────────────────────────────────────────┐
│  ╭─────────────────────╮                                        │
│  │ 🔍 Search settings  │                        [Save] [Cancel] │
│  ╰─────────────────────╯                                        │
├──────────────────────┬──────────────────────────────────────────┤
│                      │                                          │
│   ▸ Profile          │  ┌─ Profile ─────────────────────────┐   │
│   ▸ Desktop          │  │                                   │   │
│   ▸ Keys & Buttons   │  │  Doctor Name: [______________]    │   │
│   ▸ Text & Templates │  │                                   │   │
│   ▸ Alerts           │  │  ☑ Show Tooltips                  │   │
│   ▸ Report Display   │  │                                   │   │
│   ▸ Behavior         │  └───────────────────────────────────┘   │
│   ▸ RVU & Metrics    │                                          │
│   ▸ Experimental     │  ┌─ Desktop ─────────────────────────┐   │
│   ▸ Reference        │  │                                   │   │
│                      │  │  IV Report Hotkey: [Ctrl+Q]       │   │
│                      │  │                                   │   │
│                      │  │  ☑ Show Recording Indicator       │   │
│   ─────────────────  │  │    ☐ Hide when no study open      │   │
│   v3.0.8             │  │                                   │   │
│                      │  │  ☑ Auto-Stop Dictation            │   │
│                      │  │                                   │   │
│                      │  │  Audio Feedback ──────────────    │   │
│                      │  │  ☑ Start Beep  [====░░░░] 60%     │   │
│                      │  │  ☑ Stop Beep   [====░░░░] 60%     │   │
│                      │  │  Pause: [1000] ms                 │   │
│                      │  │                                   │   │
│                      │  └───────────────────────────────────┘   │
│                      │                                          │
│                      │  ┌─ Keys & Buttons ──────────────────┐   │
│                      │  │                                   │   │
│                      │  │  ...                              │   │
│                      │  │                                   │   │
│                      │  └───────────────────────────────────┘   │
│                      │                                          │
└──────────────────────┴──────────────────────────────────────────┘
```

---

## Section Reorganization

| Section | Contents |
|---------|----------|
| **Profile** | Doctor name, Show tooltips |
| **Desktop** | IV hotkey, Recording indicator, Auto-stop, Audio feedback (beeps, volumes, pause) |
| **Keys & Buttons** | Action mappings (hotkeys + mic buttons), IV Buttons studio, Window/Level keys |
| **Text & Templates** | Report font, Templates (critical, series, comparison), Macros, Pick Lists, Separate pasted items |
| **Alerts** | Critical Studies Tracker, Notification Box settings, Alert Triggers (template mismatch, gender check, stroke detection) |
| **Report Display** | Report Changes highlighting (color, alpha), Rainbow Mode (correlation), Transparent overlay, Show report after process, Show impression popup |
| **Behavior** | Scrape Mosaic, Restore focus, Scroll thresholds, Inpatient XR handling |
| **RVU & Metrics** | RVUCounter enable, Metrics selection, Goal, Database path |
| **Experimental** | Network monitor, Auto-update |
| **Reference** | AHK docs, Debug tips (read-only info box) |

---

## Implementation Architecture

### File Structure
```
UI/
├── SettingsForm.cs          # Main form - layout, search, navigation
├── SettingsSection.cs       # Base class for section panels
├── Settings/
│   ├── ProfileSection.cs
│   ├── DesktopSection.cs
│   ├── KeysButtonsSection.cs
│   ├── TextTemplatesSection.cs
│   ├── AlertsSection.cs
│   ├── ReportDisplaySection.cs
│   ├── BehaviorSection.cs
│   ├── RvuMetricsSection.cs
│   ├── ExperimentalSection.cs
│   └── ReferenceSection.cs
```

### Core Components

#### 1. SettingsForm.cs (~400 lines)
- Form setup (size, dark title bar)
- Search TextBox with filtering logic
- Navigation sidebar (custom-drawn Panel)
- Content panel (Panel with smooth scroll)
- Save/Cancel buttons
- Coordinates section visibility and scrolling

#### 2. SettingsSection.cs (~150 lines)
Base class providing:
- Dark card-style panel (rounded corners optional)
- Header with title + collapse arrow
- Auto-registers all child Label controls for tooltip
- Search matching: checks title + all label text
- Collapse/expand animation (optional)

```csharp
public abstract class SettingsSection : Panel
{
    public string Title { get; }
    public string[] SearchTerms { get; } // Computed from all labels

    public bool MatchesSearch(string query);
    public void ScrollIntoView(Panel container);
}
```

#### 3. Individual Section Classes (~150-300 lines each)
Each section class:
- Inherits SettingsSection
- Creates its controls in constructor
- Implements `LoadSettings(Configuration config)`
- Implements `SaveSettings(Configuration config)`

---

## Search Behavior

**How it works:**
1. User types in search box
2. Each section's `MatchesSearch(query)` is called
3. Sections that don't match are hidden (Visible = false)
4. Navigation items for hidden sections are grayed out
5. On search clear, all sections become visible again

**Match logic:**
- Section title contains query (case-insensitive)
- OR any label/checkbox text in that section contains query
- Entire section shows/hides together (never individual controls)

**Example:**
- Search "stroke" → Only "Alerts" section visible
- Search "beep" → Only "Desktop" section visible
- Search "RVU" → Only "RVU & Metrics" section visible

---

## Smooth Scrolling Implementation

```csharp
private void ScrollToSection(SettingsSection section)
{
    int targetY = section.Top;
    int currentY = -_contentPanel.AutoScrollPosition.Y;

    // Animate over 200ms with ease-out
    var timer = new System.Windows.Forms.Timer { Interval = 16 }; // ~60fps
    int startTime = Environment.TickCount;
    int duration = 200;

    timer.Tick += (s, e) =>
    {
        int elapsed = Environment.TickCount - startTime;
        float progress = Math.Min(1f, elapsed / (float)duration);
        float eased = 1f - (1f - progress) * (1f - progress); // ease-out quad

        int newY = (int)(currentY + (targetY - currentY) * eased);
        _contentPanel.AutoScrollPosition = new Point(0, newY);

        if (progress >= 1f) timer.Stop();
    };
    timer.Start();
}
```

---

## Tooltip Strategy

**Approach:** Standard WinForms ToolTip on label controls directly.

```csharp
// In SettingsSection base class
protected void AddSettingRow(string labelText, Control control, string tooltip)
{
    var label = new Label { Text = labelText, ... };
    Controls.Add(label);
    Controls.Add(control);

    if (!string.IsNullOrEmpty(tooltip))
    {
        _parentForm.SettingsToolTip.SetToolTip(label, tooltip);
        _parentForm.SettingsToolTip.SetToolTip(control, tooltip);
    }
}
```

**Benefits:**
- Hover the label text OR the control to see tooltip
- No question mark icons cluttering the UI
- Respects "Show Tooltips" setting (just disable ToolTip.Active)

---

## Navigation Sidebar

Custom-painted Panel with:
- Fixed width (~180px)
- Section items drawn with:
  - Icon (optional Unicode symbol or none)
  - Section name
  - Hover highlight (subtle lighter background)
  - Selected indicator (left accent bar)
- Version number at bottom
- Click handler scrolls to section

```csharp
private void DrawNavigationItem(Graphics g, int index, bool selected, bool hovered)
{
    var rect = GetItemRect(index);

    // Background
    if (selected)
        g.FillRectangle(Brushes.DimGray, rect);
    else if (hovered)
        g.FillRectangle(new SolidBrush(Color.FromArgb(50, 50, 50)), rect);

    // Selection indicator
    if (selected)
        g.FillRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)), 0, rect.Y, 3, rect.Height);

    // Text
    var section = _sections[index];
    var textColor = section.Visible ? Color.White : Color.Gray;
    g.DrawString(section.Title, Font, new SolidBrush(textColor), rect.X + 12, rect.Y + 8);
}
```

---

## Migration Strategy

1. **Create new files** alongside existing SettingsForm.cs
2. **Build incrementally** - one section at a time
3. **Test each section** before moving to next
4. **Final swap** - rename old to SettingsFormOld.cs, new to SettingsForm.cs
5. **Remove old file** once validated

---

## Estimated Line Counts

| File | Lines | Notes |
|------|-------|-------|
| SettingsForm.cs | ~400 | Main orchestration |
| SettingsSection.cs | ~150 | Base class |
| ProfileSection.cs | ~80 | Simplest section |
| DesktopSection.cs | ~250 | Beeps, indicator, hotkey |
| KeysButtonsSection.cs | ~400 | Action mappings + Button Studio |
| TextTemplatesSection.cs | ~300 | Templates, macros, pick lists |
| AlertsSection.cs | ~350 | Notification box, stroke, gender |
| ReportDisplaySection.cs | ~250 | Changes, rainbow, transparency |
| BehaviorSection.cs | ~200 | Scraping, scrolling, focus |
| RvuMetricsSection.cs | ~200 | RVU settings |
| ExperimentalSection.cs | ~100 | Network monitor |
| ReferenceSection.cs | ~50 | Static info text |
| **Total** | **~2,700** | Similar to current but organized |

---

## Key Differences from Current

| Aspect | Current | New |
|--------|---------|-----|
| Navigation | TabControl (horizontal tabs) | Sidebar (vertical list) |
| Search | None | Full-text section filtering |
| Scrolling | Tab-based, jumpy | Single pane, smooth animated |
| Tooltips | Question mark icons | Hover on labels directly |
| Code organization | Single 2900-line file | 12 focused files |
| Sections | GroupBox inside tabs | Standalone card panels |
| Visual style | Basic dark | Fluent-inspired dark |

---

## Questions Before Implementation

1. **Collapse/expand sections?** Windows 11 Settings doesn't collapse - everything scrolls. Do you want collapsible cards or just scroll-to navigation?

2. **Icons in sidebar?** Simple Unicode symbols (▸) or no icons at all (text-only like Windows Settings)?

3. **Button Studio complexity** - The IV Buttons tab has significant interactive UI (live preview, button list, editor). Keep as-is or simplify?

4. **Section order** - The proposed order above makes sense? Or different priority?

5. **Form size** - Current is 500x550. Sidebar adds width. Target ~650x600?

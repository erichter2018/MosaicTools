# Mosaic CDP Capabilities Inventory

**Generated:** 2026-03-03
**Source:** CDP probe against rp.radpair.com iframe (Mosaic reporting UI)
**Port:** Read from `%LOCALAPPDATA%\Packages\MosaicInfoHub_*\LocalState\EBWebView\DevToolsActivePort`

---

## Architecture Overview

- **App framework:** Next.js (React) + MUI (Material UI)
- **Rich text editor:** Tiptap (ProseMirror wrapper)
- **Feature flags:** LaunchDarkly (SDK key: `68c8686bb1bc780aa7a9351c`)
- **Analytics:** PostHog (`phc_yZgA0tswnVefDFt8JCLHBfE4aRDPLlC1kHtPSc9kplY`)
- **Error tracking:** Sentry (Next.js integration)
- **API backend:** `https://api-rp.radpair.com`
- **Root element:** `<div id="__next">`
- **React fiber:** NOT accessible from DOM (production build strips it)
- **Redux/MobX:** Not used (no store on window)

---

## ProseMirror Editors

Two editors exist when a study is open:

### Editor 0 — Transcript (dictation box)
- **Selector:** `document.querySelectorAll('.ProseMirror')[0]`
- **JS access:** `.editor` property on DOM element
- **Placeholder text:** "Start dictating your findings here..."
- **Node types:** paragraph, doc, text (simple — no headings, lists, formatting)
- **Marks:** highlight, rationale, uuid
- **32 plugins** including keymap handlers, paste handler, bubble menu
- **Custom extensions:** `rationale` (mark), `uuid` (mark), `placeholder`, `bubbleMenu`

### Editor 1 — Final Report (structured report)
- **Selector:** `document.querySelectorAll('.ProseMirror')[1]`
- **Node types:** paragraph, inlineParagraph, heading, draggableItem, draggableInline, listItem, section, text, orderedList, doc
- **Marks:** highlight, bold, italic, underline
- **42 plugins** including clipboard extension, section handling, draggable items
- **Custom extensions:** `section`, `draggableItem`, `draggableInline`, `inlineParagraph`, `clipboardExtension`

### Available Tiptap Commands (both editors share most)

**Content manipulation:**
- `insertContent(text)` / `insertContentAt(pos, text)` — insert text/HTML
- `setContent(content)` — replace entire editor content
- `clearContent()` — clear all content
- `deleteRange({from, to})` / `deleteSelection()` / `deleteCurrentNode()`
- `setTextSelection({from, to})` / `selectAll()`

**Formatting (Editor 1 only):**
- `setBold()` / `toggleBold()` / `unsetBold()`
- `setItalic()` / `toggleItalic()` / `unsetItalic()`
- `setUnderline()` / `toggleUnderline()` / `unsetUnderline()`
- `setHeading({level})` / `toggleHeading({level})`
- `toggleOrderedList()` / `wrapInList()` / `liftListItem()` / `sinkListItem()`
- `setHighlight()` / `toggleHighlight()` / `unsetHighlight()`

**Cursor/selection:**
- `focus('start'|'end'|pos)` / `blur()`
- `setNodeSelection(pos)` / `selectParentNode()`
- `scrollIntoView()`

**History:**
- `undo()` / `redo()` — 100-depth history, 500ms grouping delay

**Structural:**
- `setNode(type)` / `setParagraph()` / `toggleNode(type)`
- `splitBlock()` / `joinUp()` / `joinDown()` / `lift()` / `wrapIn(type)`

### State & Document Access
- `editor.state.doc` — ProseMirror document tree
- `editor.state.doc.textBetween(from, to)` — extract text range
- `editor.state.doc.content.size` — total document size
- `editor.state.doc.descendants(callback)` — walk all nodes
- `editor.state.selection` — current selection `{from, to}`
- `editor.getHTML()` — get content as HTML string
- `editor.getJSON()` — get content as structured JSON
- `editor.isEditable` — editability state

### Event Hooks
- `onUpdate` — fires after content changes
- `onSelectionUpdate` — fires on cursor/selection change
- `onTransaction` — fires on every ProseMirror transaction
- `onFocus` / `onBlur` — focus events
- `onCreate` / `onDestroy` — lifecycle

### Keyboard Shortcuts (registered via extensions)
**Editor 0:** paragraph, keymap, history, highlight
**Editor 1:** paragraph, keymap, draggableItem, listItem, section, bold, history, italic, underline, orderedList

---

## RadPair REST API Endpoints

All calls go to `https://api-rp.radpair.com`. Auth via token (verified at `/users/tokens/verify`).

### User & Config
| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/users/tokens/verify?get_user_info=true` | GET | Auth verification + user info |
| `/preferences/all` | GET | **All user preferences** |
| `/permissions/?max_results=0` | GET | User permissions |
| `/orgs/1/settings/?max_results=0` | GET | Organization settings |

### Reports
| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/reports/?max_results=20` | GET | Report list |
| `/reports/{id}` | GET | Single report data |
| `/reports/{id}/opened` | POST? | Notify report opened |
| `/reports/{id}/update` | POST? | Save report changes |
| `/reports/{id}/diff-check` | POST? | Check for report changes/conflicts |
| `/v2/reports/processing/process/{id}?normal=true` | POST? | **Process Report** (4.9s response!) |
| `/v2/reports/templates/?study_id=X&org_id=Y` | GET | **Get templates for study type** |
| `/v2/reports/studies/hybrid_search?query=&max_most_used=10` | GET | **Study type search** (used by autocomplete) |

### Macros
| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/macros/available` | GET | **GET ALL AVAILABLE MACROS** |

### Quality Checks
| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/reports/check/laterality` | POST? | Laterality mismatch check |
| `/reports/check/tarsal-carpal` | POST? | Tarsal/carpal bone check |

### External Services
| Service | Purpose |
|---------|---------|
| LaunchDarkly (`app.launchdarkly.com`) | Feature flags |
| PostHog (`us.i.posthog.com`) | Analytics, session replay |
| Sentry (`ingest.us.sentry.io`) | Error reporting |
| Google Tag Manager (`googletagmanager.com`) | Analytics |

---

## Mosaic Macros — Key Finding

**The `/macros/available` endpoint is called when a study is opened.** This means:
1. Macro definitions are fetched from the RadPair API, not stored locally
2. We could potentially call this endpoint directly to get all macro names/definitions
3. Macro expansion is likely handled SERVER-SIDE (explains why typing trigger text didn't work — it's speech-triggered, not text-triggered)
4. The `macrosPanel` in localStorage's `accordion-state-storage` confirms there's a macros UI panel in Mosaic's sidebar

**What we know about Mosaic macro triggers:**
- NOT triggered by typing keyword text into the editor
- Triggered by VOICE COMMAND (speech recognition detects the macro name)
- Definitions fetched from `/macros/available` API
- The `handleTextInput` plugin on editor_0 (key: `plugin$19`) intercepts text input — this could be where macro detection happens, but it seems to be speech-driven

**Next steps to investigate macros:**
1. Intercept the response from `/macros/available` to see the macro format/schema
2. Check if macros can be triggered via the API (POST endpoint?)
3. The `handleTextInput` plugin on editor_0 might be the expansion mechanism — could hook into it

---

## UI Controls (with study open)

### Named Buttons
| aria-label / text | Class hint | Notes |
|-------------------|-----------|-------|
| "My Configuration" | IconButton-sizeLarge | Mosaic settings/config panel |
| "Share Feedback" | radpair-feedback-appbar-button | Feedback form |
| "Help & Support" | IconButton-sizeLarge | Help docs |
| "Clear" | MuiAutocomplete-clearIndicator | Clear autocomplete (study type?) |
| "Open" | MuiAutocomplete-popupIndicator | Open autocomplete dropdown (x2) |
| "ACTIONS" | MuiButton-outlinedWarning | **Actions menu** (orange outlined) |
| "PROCESS REPORT" | MuiButton-containedPrimary | Process Report button |
| "CREATE IMPRESSION" | MuiButton-outlinedInfo | Create Impression button |
| "SIGN ADDENDUM" | MuiButton-outlinedSuccess | Sign button (shows "ADDENDUM" for this study) |
| "Open documentation in new tab" | IconButton-sizeSmall | Doc link |
| (Record button) | MuiButton-containedRecord | **Dictation record button** (custom "Record" color) |

### Toggle Buttons (toolbar)
- ~15 MuiToggleButton elements (formatting toolbar — bold, italic, underline, lists, etc.)
- Some disabled (context-dependent)
- One has `Mui-selected` (currently active formatting)

### Inputs
- Study type combobox: `[role="combobox"]` with `#studies-search` input
- Template selector: `#templates-select` input
- Two "Open" popup indicator buttons (one per autocomplete)

---

## DOM Structure (high-level)

```
body
└── #__next
    ├── MuiBackdrop (loading overlays x2)
    ├── radpair-app-bar (top bar)
    │   └── MuiToolbar
    │       ├── Logo/link
    │       └── MuiStack (My Config, Feedback, Help, User buttons)
    ├── MuiBox (main content area)
    │   ├── react-joyride (tour/walkthrough system)
    │   ├── radpair-reports-drawer (left sidebar - study list)
    │   │   └── MuiList (study items with aria-labels like "XR Chest")
    │   └── main (reporting area)
    │       ├── MuiToolbar (main toolbar)
    │       └── Toastify (toast notifications)
    └── NEXT-ROUTE-ANNOUNCER (a11y route announcements)
```

---

## localStorage State

| Key | Content |
|-----|---------|
| `accordion-state-storage` | Panel open/close state: `{labelPanel, macrosPanel, reportInfoPanel, clinicalDataPanel}` |
| `processing-tasks-storage` | Background processing tasks |
| `userInfo` | User profile (id: 4635, email, name, org: RADPARTNERS, role: user) |
| `sidebar-state-storage` | Sidebar collapsed state |
| `tour-status-storage` | UI tour completion status |
| `ld:*` | LaunchDarkly feature flag state |
| `ph_*_posthog` | PostHog analytics state |

---

## What's NOT in the Web Layer

- **Dragon/PowerScribe dictation** — handled outside the iframe (native app layer). No voice/dictation DOM elements, no WebSocket for speech, no MediaRecorder usage detected. Speech recognition integration happens at the MosaicInfoHub (SlimHub) native app level, not in the web UI.
- **No Service Worker** — app doesn't work offline
- **No nested iframes** — the reporting iframe is the innermost level
- **No Redux/state management store** on window — state is component-local or context-based

---

## Actionable Opportunities

### 1. Fetch Macro Definitions via API
Call `https://api-rp.radpair.com/macros/available` with the user's auth token to get all macro names, trigger words, and content. Token can be extracted from localStorage `userInfo` or intercepted from `/users/tokens/verify`.

### 2. Use `editor.getJSON()` for Structured Report Parsing
Instead of walking DOM nodes, `editor.getJSON()` returns the full ProseMirror document as a structured JSON tree. More reliable than text extraction for finding sections, numbered lists, headings.

### 3. Programmatic Undo (`editor.commands.undo()`)
If an insertion goes wrong, can undo it cleanly through the editor's history system.

### 4. `setContent()` for Full Report Replacement
Could replace the entire report content at once (e.g., template swap) instead of select-all + paste.

### 5. Laterality/Quality Check APIs
`/reports/check/laterality` and `/reports/check/tarsal-carpal` — could run these checks proactively before Process Report to warn about errors.

### 6. Template API
`/v2/reports/templates/?study_id=X&org_id=Y` — could fetch the full list of templates for a study type, potentially useful for template correction.

### 7. Study Search API
`/v2/reports/studies/hybrid_search?query=&max_most_used=10` — could show most-used study types or search programmatically.

### 8. Click "ACTIONS" Button
The ACTIONS button (orange outlined, `MuiButton-outlinedWarning`) likely opens a dropdown with additional actions. Could click it via CDP and enumerate what's inside.

### 9. Reports Drawer / Study List
The left sidebar (`radpair-reports-drawer`) shows all studies for the patient. SPANs have `aria-label` with study type names. Could read all assigned studies.

### 10. Network Interception via CDP
CDP supports `Network.enable` to see all request/response data in real-time. Could intercept:
- `/macros/available` response to get macro definitions
- `/reports/{id}` response for full report data
- `/v2/reports/processing/process/{id}` to know when processing completes

---

## Probe Script

The probe script that generated this data is at `cdp_full_inventory.ps1` in the project root. Run it when MosaicTools is NOT connected (kill MT first). Requires a study to be open in Mosaic for full results.

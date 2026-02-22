# Mosaic Native STT Macro Investigation Notes (Sanitized)

Date: 2026-02-22
Repo: `C:\Users\erik.richter\Desktop\MosaicTools`

## Goal
Understand whether MosaicTools custom STT can trigger Mosaic's built-in voice macros, and if so, how.

## Initial conclusion (important)
Custom STT in MT currently inserts text after recognition, so it bypasses Mosaic's native STT command parser.

- MT custom STT path in `MosaicToolsCSharp\UI\ActionController.cs` receives final transcript and inserts text via paste/SendInput.
- Mosaic native macros are not triggered by typed/pasted text.
- Therefore, custom STT cannot trigger native Mosaic voice macros through the current direct-paste path.

## Key code findings in MT
- `MosaicToolsCSharp\UI\ActionController.cs`
  - Custom STT subscription and direct insert path
  - Final transcripts are cleaned and inserted as plain text
- `MosaicToolsCSharp\Services\SttService.cs`
  - Streams audio and relays transcript events only (no macro/command parsing)
- `MosaicToolsCSharp\Services\NativeWindows.cs`
  - `SendUnicodeText()` exists (SendInput text insertion), but this still does not trigger Mosaic native macros

## Recon tools added to this repo during investigation
### 1) `MosaicSttRecon`
Purpose: compare normal phrase vs macro phrase activity (processes, TCP connections, named pipes)

Files:
- `MosaicSttRecon\Program.cs`
- `MosaicSttRecon\README.md`

What it found:
- `MosaicInfoHub` runs a local Kestrel server on `127.0.0.1:5002`
- Mosaic package includes ASP.NET Core / WebAPI / Swashbuckle components
- No obvious macro-specific TCP/pipe difference from coarse snapshots (mostly noise / existing connections)

### 2) `ProcFreeze`
Purpose: suspend/resume a process (and later added `trap` mode) to try to catch transient Clario XML files

Files:
- `ProcFreeze\Program.cs`

Result:
- Suspending `MosaicInfoHub` breaks dictation (important clue: it is in the active dictation pipeline)
- File watcher/trap approach did not reliably catch transient XML payloads for macro investigation

## Major breakthrough: local MosaicInfoHub / SlimHub findings
Installed app path found:
- `C:\Program Files\WindowsApps\MosaicInfoHub_2.0.4.0_x64__yzjagrpanxm4c\SlimHub\`

Important config files discovered:
- `apisettings.json`
- `Configuration\appsettings.json`

Confirmed in `apisettings.json`:
- Kestrel endpoint on `http://localhost:5002`

Important logs discovered:
- `C:\ProgramData\Mosaic\SlimHub-YYYY-MM-DD.log`

Log observations:
- `SlimHub` logs are verbose and include WebView postMessage traffic and Clario file-drop events
- Clario dictation file watcher monitors:
  - `C:\MModal\FluencyForImaging\Reporting\XML\IN\`
- Common Clario message types observed in logs:
  - `SetTokenValue` (most common; likely dictation token path)
  - `OpenReport`
  - `Show`
  - `Login`
  - `Exit`

## DevTools investigation (no admin needed)
The `DevTools - localhost/index.html` window was used to inspect network traffic.

### Important observations
- WebSocket rows with names like `v1?...join_request...` are connection handshakes, not macro payloads
- WebSocket message frames appeared binary
- More useful data was in XHR/Fetch requests, especially:
  - `dictation-session`
  - `process-and-format-transcript`
  - `update`

## Critical finding: macro expansion API behavior
Endpoint observed:
- `POST https://api-rp.radpair.com/reports/{reportId}/process-and-format-transcript`

This endpoint accepts plain JSON transcript text, e.g.:
- `new_text`
- `fix_transcript`
- `cursor_start`
- `cursor_end`
- `original_text`
- `mode`

### For normal speech
Response returns `processed_text` with normalized text (example: `" testing 123 "`).

### For macro phrase (example)
Response returned expanded macro content in `processed_text`, not just the spoken phrase.

This proves:
- Mosaic/RadPair macro expansion is accessible via a text-processing API step
- Macro behavior is not strictly tied to audio input at the point of expansion

## Confirmed auth requirement
Unauthenticated manual call to `process-and-format-transcript` returned:
- `401 Unauthorized`

This proves authentication is required for this endpoint.

## What is likely required for MT to call the macro expansion API
High confidence:
- Valid authenticated session cookies (RadPair/Mosaic web session)
- Current `reportId` (route is report-scoped)

Possibly required / still to confirm:
- `x-session-id`

Likely *not* strictly required (usually):
- Most browser fingerprint headers (`sec-*`, user-agent) if auth/session is valid

## What would be required to make MT custom STT trigger Mosaic macro expansions
Conceptually:
1. MT receives final text from custom STT
2. MT calls `process-and-format-transcript` using active authenticated Mosaic/RadPair session context
3. MT reads `processed_text`
4. MT inserts `processed_text` into Mosaic transcript box (or report editor)

This would preserve Mosaic macro expansions while using custom STT.

## Main engineering blocker now
MT does not currently have access to live RadPair/Mosaic web session context:
- auth cookies
- current `reportId`
- possibly `x-session-id`

That is the remaining hard problem, not macro recognition.

## Risk / policy notes (important)
This path is technically feasible but unofficial.
Potential concerns:
- Internal API use not intended for third-party tooling
- Handling auth cookies/session data in MT
- PHI/security/logging risks if payloads are persisted
- Vendor/workflow breakage risk if endpoint contracts change

Recommendation if continuing:
- Treat as proof-of-concept first
- Do not persist auth tokens/cookies on disk
- Disable payload logging
- Feature-flag it
- Consider discussing with local IT/security / governance before production use

## Sanitization note
Live auth cookies / tokens / session IDs were observed during investigation and should be considered sensitive.
Do not store them in this repo. Rotate/re-authenticate if previously shared.

## Useful commands (sanitized)
### Run recon tool
```powershell
& 'C:\Users\erik.richter\Desktop\dotnet\dotnet.exe' run --project .\MosaicSttRecon\MosaicSttRecon.csproj -c Release
```

### Tail SlimHub log
```powershell
Get-Content (Get-ChildItem 'C:\ProgramData\Mosaic' -Filter 'SlimHub-*.log' | Sort-Object LastWriteTime -Descending |
  Select-Object -First 1 -ExpandProperty FullName) -Wait -Tail 80
```

### Test unauthenticated formatter endpoint (expected 401)
```powershell
try {
  Invoke-RestMethod \
    -Uri 'https://api-rp.radpair.com/reports/<reportId>/process-and-format-transcript' \
    -Method POST \
    -ContentType 'application/json' \
    -Body '{"new_text":"Macro differential.","fix_transcript":false,"cursor_start":0,"cursor_end":0,"original_text":"","mode":"improved"}'
} catch {
  $_.Exception.Response.StatusCode.value__
  $_.Exception.Message
}
```

## If continuing later: recommended next technical steps
1. Determine a safe way for MT to obtain current `reportId` and authenticated web session context (cookies / maybe `x-session-id`).
2. Build a standalone proof-of-concept caller for `process-and-format-transcript` (manual/session-injected config, no persistence).
3. Integrate into MT custom STT path as an optional feature:
   - final transcript -> formatter API -> `processed_text` -> insert
   - fallback to raw transcript on failure.
4. Add strict logging redaction and local-only ephemeral token handling.

# Mosaic STT Recon (Guided)

Purpose: help compare what changes on your machine when you say a normal phrase vs a Mosaic voice macro phrase.

## What it captures
- Mosaic-related processes (including child/helper processes)
- TCP connections owned by those processes (`netstat` snapshots)
- Named pipe names (system-wide snapshot, best effort)
- A plain-English hint section in the report

## What it does NOT capture
- Encrypted message contents
- Exact WebSocket payloads
- Internal Mosaic function calls

## How to run
1. Open Mosaic and load a study.
2. Open PowerShell in this repo (`C:\Users\erik.richter\Desktop\MosaicTools`).
3. Run:

```powershell
& 'C:\Users\erik.richter\Desktop\dotnet\dotnet.exe' run --project .\MosaicSttRecon\MosaicSttRecon.csproj -c Release
```

4. When prompted for Phase 1, press Enter and immediately dictate a normal phrase (not a macro).
5. When prompted for Phase 2, press Enter and immediately say a Mosaic macro phrase.
6. Wait for both 10-second captures to finish.
7. Open the report written to your Desktop (file name starts with `mosaic_stt_recon_`).

## How to read the report (first pass)
- If you see a localhost connection or suspicious pipe that appears only during the macro phrase, that is a strong lead.
- If the same connections are present in both phases, the macro command may be inside existing traffic (likely WebSocket payloads).
- If there are no obvious differences, the macro parser may be in-process and we should move to API monitoring/hooking.

## What to send back to me
- The report file contents (or just the sections for "Differences" and "Interpretation Hints")
- Whether the macro phrase definitely fired in Mosaic during Phase 2
- The exact macro phrase you used (or a sanitized version)

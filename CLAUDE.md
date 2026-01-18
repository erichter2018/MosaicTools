# MosaicTools Project

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
```powershell
taskkill /IM MosaicTools.exe /F 2>$null; Start-Sleep -Seconds 1; c:\Users\erik.richter\Desktop\dotnet\dotnet.exe publish -c Release -r win-x64 --self-contained; Start-Process "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe"
```

### Quick Debug Build (syntax checking only)
```powershell
c:\Users\erik.richter\Desktop\dotnet\dotnet.exe build MosaicTools.csproj
```

### Common Issues
1. **"File is being used by another process"** - Kill MosaicTools.exe first, wait a second
2. **"Project file does not exist"** - Use `MosaicTools.csproj`, not `MosaicToolsCSharp.csproj`
3. **"dotnet is not recognized"** - Use full path: `c:\Users\erik.richter\Desktop\dotnet\dotnet.exe`

---

## GitHub Repository

- **URL:** https://github.com/erichter2018/MosaicTools
- **Account:** erichter@gmail.com

---

## Future Feature: Auto-Update System

**Status:** Not implemented yet - discussed January 2026

### Recommended Approach: GitHub Releases + Simple Downloader

Since the repo is already on GitHub, we can use the GitHub Releases API directly:
- `https://api.github.com/repos/erichter2018/MosaicTools/releases/latest`
- No need for separate hosting of `latest.json`

Since MosaicTools is distributed as a single self-contained exe, the recommended approach is:

1. **Version check mechanism:**
   - Host a `latest.json` file (on GitHub releases, S3, or simple web server) containing:
     ```json
     {"version": "2.1", "url": "https://..../MosaicTools.exe", "notes": "Bug fixes..."}
     ```
   - On app startup, fetch this JSON and compare with current version
   - If newer version exists, notify user

2. **Update delivery options (choose one):**
   - **Simple:** Show "Update available" toast that opens browser to download page - user manually replaces exe
   - **Better:** Download new exe in background to temp location, then use batch script replacement

3. **Batch script replacement method:**
   - App downloads new exe to temp folder
   - App creates and launches a small batch script:
     ```batch
     timeout 2
     copy /y "temp\MosaicTools_new.exe" "MosaicTools.exe"
     start MosaicTools.exe
     del updater.bat
     ```
   - App exits, batch script waits, replaces exe, restarts app

4. **Alternative libraries considered:**
   - **Squirrel.Windows** - Battle-tested (Slack/Discord use it), but changes deployment model to installer-based
   - **ClickOnce** - Built into .NET but feels dated, limited customization
   - Both overkill for a single-exe distribution

### Implementation Notes
- Version is already set in `MosaicTools.csproj` (currently 2.1.0)
- Version displayed in startup toast and Settings form title
- GitHub API can be used to check releases (generous rate limits)
- Consider checking for updates only once per day to avoid annoyance

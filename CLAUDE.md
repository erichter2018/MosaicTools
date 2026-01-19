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
- **Visibility:** Public (required for auto-update to work)

### gh CLI Location
```
"C:\Users\erik.richter\Desktop\GH CLI\gh.exe"
```

---

## Auto-Update System

The app auto-updates via GitHub Releases using a rename trick (no batch files, no installer, no admin required).

### How it works
1. On startup, checks `https://api.github.com/repos/erichter2018/MosaicTools/releases/latest`
2. If newer version found, downloads exe to `MosaicTools_new.exe`
3. Renames running exe to `MosaicTools_old.exe` (Windows allows renaming running files)
4. Renames new exe to `MosaicTools.exe`
5. Shows toast with "Restart Now" button
6. On next startup, deletes `_old.exe`

### Publishing a new release

When user says "create a release" or "publish release vX.X":

1. **Update version** in `MosaicToolsCSharp/MosaicTools.csproj`:
   ```xml
   <Version>2.2.0</Version>
   ```

2. **Commit and push** the version change:
   ```bash
   git add -A && git commit -m "Bump version to 2.2.0" && git push
   ```

3. **Build** the release exe:
   ```bash
   c:\Users\erik.richter\Desktop\dotnet\dotnet.exe publish MosaicTools.csproj -c Release -r win-x64 --self-contained
   ```

4. **Create the GitHub release** using gh CLI:
   ```bash
   "C:\Users\erik.richter\Desktop\GH CLI\gh.exe" release create v2.2.0 "C:\Users\erik.richter\Desktop\MosaicTools\MosaicToolsCSharp\bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MosaicTools.exe" --title "v2.2.0" --notes "Release notes here"
   ```

### Settings
- **Auto-update** checkbox in General tab (ON by default)
- **Check for Updates** button for manual checks

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MosaicTools.Services;

/// <summary>
/// Handles auto-update via GitHub Releases.
/// Downloads zip files to avoid corporate security blocking exe downloads.
/// </summary>
public class UpdateService
{
    private const string GitHubApiUrl = "https://api.github.com/repos/erichter2018/MosaicTools/releases/latest";
    private const string GitHubAllReleasesUrl = "https://api.github.com/repos/erichter2018/MosaicTools/releases";
    private static readonly HttpClient _httpClient;
    private static readonly SemaphoreSlim _updateLock = new(1, 1);

    public string? LatestVersion { get; private set; }
    public string? DownloadUrl { get; private set; }
    public string? ReleaseNotes { get; private set; }
    public bool UpdateReady { get; private set; }

    // Retry constants for file operations (OneDrive sync, antivirus locks)
    private const int FileOpMaxRetries = 5;
    private const int FileOpRetryDelayMs = 500;

    static UpdateService()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.Add("User-Agent", "MosaicTools-AutoUpdater");
    }

    #region WinVerifyTrust (Authenticode)

    private const uint WTD_UI_NONE = 2;
    private const uint WTD_CHOICE_FILE = 1;
    private const uint WTD_REVOKE_NONE = 0;
    private const uint WTD_STATEACTION_VERIFY = 1;
    private static readonly Guid WINTRUST_ACTION_GENERIC_VERIFY_V2 =
        new("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WINTRUST_FILE_INFO
    {
        public uint cbStruct;
        public IntPtr pcwszFilePath;
        public IntPtr hFile;
        public IntPtr pgKnownSubject;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WINTRUST_DATA
    {
        public uint cbStruct;
        public IntPtr pPolicyCallbackData;
        public IntPtr pSIPClientData;
        public uint dwUIChoice;
        public uint fdwRevocationChecks;
        public uint dwUnionChoice;
        public IntPtr pFile;
        public uint dwStateAction;
        public IntPtr hWVTStateData;
        public IntPtr pwszURLReference;
        public uint dwProvFlags;
        public uint dwUIContext;
        public IntPtr pSignatureSettings;
    }

    [DllImport("wintrust.dll", CharSet = CharSet.Unicode)]
    private static extern int WinVerifyTrust(IntPtr hwnd, ref Guid pgActionID, ref WINTRUST_DATA pWVTData);

    /// <summary>
    /// Verify Authenticode signature on a file. Returns true if validly signed.
    /// Returns true (skip check) if WinVerifyTrust is unavailable.
    /// </summary>
    private static bool VerifyAuthenticode(string filePath)
    {
        try
        {
            var filePathPtr = Marshal.StringToCoTaskMemUni(filePath);
            try
            {
                var fileInfo = new WINTRUST_FILE_INFO
                {
                    cbStruct = (uint)Marshal.SizeOf<WINTRUST_FILE_INFO>(),
                    pcwszFilePath = filePathPtr,
                    hFile = IntPtr.Zero,
                    pgKnownSubject = IntPtr.Zero
                };

                var fileInfoPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf<WINTRUST_FILE_INFO>());
                try
                {
                    Marshal.StructureToPtr(fileInfo, fileInfoPtr, false);

                    var trustData = new WINTRUST_DATA
                    {
                        cbStruct = (uint)Marshal.SizeOf<WINTRUST_DATA>(),
                        dwUIChoice = WTD_UI_NONE,
                        fdwRevocationChecks = WTD_REVOKE_NONE,
                        dwUnionChoice = WTD_CHOICE_FILE,
                        pFile = fileInfoPtr,
                        dwStateAction = WTD_STATEACTION_VERIFY,
                    };

                    var actionId = WINTRUST_ACTION_GENERIC_VERIFY_V2;
                    int result = WinVerifyTrust(IntPtr.Zero, ref actionId, ref trustData);
                    return result == 0; // 0 = SUCCESS
                }
                finally
                {
                    Marshal.FreeCoTaskMem(fileInfoPtr);
                }
            }
            finally
            {
                Marshal.FreeCoTaskMem(filePathPtr);
            }
        }
        catch (Exception ex)
        {
            // If signature check itself fails (missing DLL, marshalling error),
            // log but allow the update — don't block updates on machines where
            // WinVerifyTrust is broken
            Logger.Trace($"Authenticode check failed (allowing update): {ex.Message}");
            return true;
        }
    }

    #endregion

    #region MOTW Zone.Identifier

    /// <summary>
    /// Remove Mark of the Web Zone.Identifier alternate data stream from a file.
    /// This prevents SmartScreen/security warnings when launching the downloaded exe.
    /// </summary>
    private static void RemoveZoneIdentifier(string filePath)
    {
        try
        {
            // Zone.Identifier is stored as an NTFS alternate data stream
            var adsPath = filePath + ":Zone.Identifier";
            // DeleteFile on the ADS path removes just the stream, not the file
            DeleteFile(adsPath);
        }
        catch
        {
            // Non-critical — may fail on non-NTFS or if no ADS exists
        }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteFile(string lpFileName);

    #endregion

    #region Resilient File Operations

    /// <summary>
    /// File.Move with retry logic for OneDrive sync locks and antivirus scanning.
    /// </summary>
    private static void MoveFileWithRetry(string source, string dest)
    {
        for (int i = 0; i < FileOpMaxRetries; i++)
        {
            try
            {
                File.Move(source, dest);
                return;
            }
            catch (IOException) when (i < FileOpMaxRetries - 1)
            {
                Thread.Sleep(FileOpRetryDelayMs);
            }
        }
    }

    /// <summary>
    /// File.Delete with retry logic for OneDrive sync locks and antivirus scanning.
    /// </summary>
    private static bool DeleteFileWithRetry(string path)
    {
        for (int i = 0; i < FileOpMaxRetries; i++)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
                return true;
            }
            catch (IOException) when (i < FileOpMaxRetries - 1)
            {
                Thread.Sleep(FileOpRetryDelayMs);
            }
            catch
            {
                return false;
            }
        }
        return false;
    }

    #endregion

    /// <summary>
    /// Get the current app version.
    /// </summary>
    public static Version GetCurrentVersion()
    {
        return Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0);
    }

    /// <summary>
    /// Check GitHub for a newer release.
    /// Returns true if a newer version is available.
    /// Includes one automatic retry for transient network failures.
    /// </summary>
    public async Task<bool> CheckForUpdateAsync()
    {
        // Reset stale state from any previous check
        LatestVersion = null;
        DownloadUrl = null;
        ReleaseNotes = null;

        for (int attempt = 0; attempt < 2; attempt++)
        {
            try
            {
                if (attempt > 0)
                {
                    Logger.Trace("Retrying update check...");
                    await Task.Delay(2000);
                }

                Logger.Trace("Checking for updates...");

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                using var response = await _httpClient.GetAsync(GitHubApiUrl, cts.Token);
                if (!response.IsSuccessStatusCode)
                {
                    Logger.Trace($"GitHub API returned {response.StatusCode}");
                    if (attempt == 0 && IsTransientHttpError(response.StatusCode)) continue;
                    return false;
                }

                var release = await response.Content.ReadFromJsonAsync<GitHubRelease>(cancellationToken: cts.Token);
                if (release == null)
                {
                    Logger.Trace("Failed to parse GitHub release");
                    return false;
                }

                var tagVersion = release.TagName?.TrimStart('v', 'V') ?? "0.0";
                if (!Version.TryParse(NormalizeVersion(tagVersion), out var latestVersion))
                {
                    Logger.Trace($"Failed to parse version from tag: {release.TagName}");
                    return false;
                }

                var currentVersion = GetCurrentVersion();
                LatestVersion = latestVersion.ToString();
                ReleaseNotes = release.Body;

                Logger.Trace($"Current: {currentVersion}, Latest: {latestVersion}");

                if (latestVersion > currentVersion)
                {
                    var zipAsset = release.Assets?.Find(a =>
                        a.Name?.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) == true);
                    var exeAsset = release.Assets?.Find(a =>
                        a.Name?.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) == true);
                    var asset = zipAsset ?? exeAsset;

                    if (asset?.BrowserDownloadUrl != null)
                    {
                        DownloadUrl = asset.BrowserDownloadUrl;
                        Logger.Trace($"Update available: {DownloadUrl}");
                        return true;
                    }
                    else
                    {
                        Logger.Trace("No zip or exe asset found in release");
                    }
                }

                return false;
            }
            catch (Exception ex) when (attempt == 0 && IsTransientException(ex))
            {
                Logger.Trace($"Update check failed (will retry): {ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.Trace($"Update check failed: {ex.Message}");
                return false;
            }
        }
        return false;
    }

    /// <summary>
    /// Download the update and prepare it for installation.
    /// Supports both zip files (preferred) and exe files (legacy).
    /// Returns true if ready to restart.
    /// </summary>
    public async Task<bool> DownloadUpdateAsync(Action<int>? progressCallback = null)
    {
        if (string.IsNullOrEmpty(DownloadUrl))
        {
            Logger.Trace("No download URL available");
            return false;
        }

        if (!await _updateLock.WaitAsync(0))
        {
            Logger.Trace("Update already in progress");
            return false;
        }

        try
        {
            var result = await DownloadAndSwapAsync(DownloadUrl, progressCallback);
            if (result) UpdateReady = true;
            return result;
        }
        finally
        {
            _updateLock.Release();
        }
    }

    /// <summary>
    /// Fetch all releases from GitHub, excluding the current version.
    /// Returns list sorted newest-first with download URLs resolved.
    /// </summary>
    public async Task<List<(string Tag, string DownloadUrl)>> GetAllReleasesAsync()
    {
        var results = new List<(string Tag, string DownloadUrl)>();
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            using var response = await _httpClient.GetAsync(GitHubAllReleasesUrl, cts.Token);
            if (!response.IsSuccessStatusCode) return results;

            var releases = await response.Content.ReadFromJsonAsync<List<GitHubRelease>>(cancellationToken: cts.Token);
            if (releases == null) return results;

            var currentVersion = GetCurrentVersion();

            foreach (var release in releases)
            {
                var tagVersion = release.TagName?.TrimStart('v', 'V') ?? "0.0";
                if (!Version.TryParse(NormalizeVersion(tagVersion), out var ver)) continue;
                if (ver == currentVersion) continue;

                var zipAsset = release.Assets?.Find(a =>
                    a.Name?.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) == true);
                var exeAsset = release.Assets?.Find(a =>
                    a.Name?.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) == true);
                var asset = zipAsset ?? exeAsset;

                if (asset?.BrowserDownloadUrl != null)
                    results.Add((release.TagName!, asset.BrowserDownloadUrl));
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Failed to fetch releases: {ex.Message}");
        }
        return results;
    }

    /// <summary>
    /// Download and install a specific version from a given URL.
    /// Same rename-trick logic as DownloadUpdateAsync but with an explicit URL.
    /// </summary>
    public async Task<bool> DownloadAndInstallVersionAsync(string downloadUrl)
    {
        if (!await _updateLock.WaitAsync(0))
        {
            Logger.Trace("Update already in progress");
            return false;
        }

        try
        {
            return await DownloadAndSwapAsync(downloadUrl);
        }
        finally
        {
            _updateLock.Release();
        }
    }

    /// <summary>
    /// Core download + extract + verify + swap logic shared by all update paths.
    /// </summary>
    private async Task<bool> DownloadAndSwapAsync(string downloadUrl, Action<int>? progressCallback = null)
    {
        var exePath = Application.ExecutablePath;
        var exeDir = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;
        var newExePath = Path.Combine(exeDir, "MosaicTools_new.exe");
        var oldExePath = Path.Combine(exeDir, "MosaicTools_old.exe");
        var isZipDownload = downloadUrl.EndsWith(".zip", StringComparison.OrdinalIgnoreCase);
        var downloadPath = isZipDownload
            ? Path.Combine(exeDir, "MosaicTools_update.zip")
            : newExePath;

        try
        {
            Logger.Trace($"Downloading update from {downloadUrl} (zip={isZipDownload})");

            // Download file (5 minute timeout for large files)
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
            using var response = await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead, cts.Token);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength ?? -1;
            var downloadedBytes = 0L;

            await using var contentStream = await response.Content.ReadAsStreamAsync(cts.Token);
            await using var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None);

            var buffer = new byte[81920];
            int bytesRead;

            while ((bytesRead = await contentStream.ReadAsync(buffer, cts.Token)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead), cts.Token);
                downloadedBytes += bytesRead;

                if (totalBytes > 0)
                {
                    var progress = (int)(downloadedBytes * 100 / totalBytes);
                    progressCallback?.Invoke(progress);
                }
            }

            fileStream.Close();
            Logger.Trace($"Download complete: {downloadedBytes} bytes");

            // Extract exe from zip if needed
            if (isZipDownload)
            {
                Logger.Trace("Extracting exe from zip...");

                try
                {
                    using var archive = ZipFile.OpenRead(downloadPath);
                    var exeEntry = archive.Entries.FirstOrDefault(e =>
                        e.Name.Equals("MosaicTools.exe", StringComparison.OrdinalIgnoreCase));

                    if (exeEntry == null)
                    {
                        Logger.Trace("No MosaicTools.exe found in zip");
                        DeleteFileWithRetry(downloadPath);
                        return false;
                    }

                    exeEntry.ExtractToFile(newExePath, overwrite: true);
                    Logger.Trace($"Extracted {exeEntry.Name} ({exeEntry.Length} bytes)");
                }
                finally
                {
                    DeleteFileWithRetry(downloadPath);
                }
            }

            // Remove MOTW Zone.Identifier from extracted/downloaded exe
            RemoveZoneIdentifier(newExePath);

            // Verify the download is a valid exe (basic size check)
            var fileInfo = new FileInfo(newExePath);
            if (fileInfo.Length < 100000)
            {
                Logger.Trace("Downloaded file too small, likely invalid");
                DeleteFileWithRetry(newExePath);
                return false;
            }

            // Verify Authenticode signature (skip unsigned dev builds gracefully)
            if (!VerifyAuthenticode(newExePath))
            {
                Logger.Trace("WARNING: Downloaded exe failed Authenticode verification — blocking update");
                DeleteFileWithRetry(newExePath);
                return false;
            }

            // Ensure _old.exe slot is free (retry for OneDrive/AV locks)
            if (File.Exists(oldExePath))
            {
                if (!DeleteFileWithRetry(oldExePath))
                {
                    Logger.Trace("Cannot delete old exe (locked) — aborting update");
                    DeleteFileWithRetry(newExePath);
                    return false;
                }
            }

            // Rename current exe to _old (Windows allows renaming a running exe)
            MoveFileWithRetry(exePath, oldExePath);

            // Rename new exe to current name — rollback first rename on failure
            try
            {
                MoveFileWithRetry(newExePath, exePath);
            }
            catch
            {
                // Rollback: restore original exe name so the app isn't left broken
                try { MoveFileWithRetry(oldExePath, exePath); } catch { }
                throw;
            }

            // Remove MOTW from the final exe path too (belt and suspenders)
            RemoveZoneIdentifier(exePath);

            Logger.Trace("Update files ready, restart required");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Download/swap failed: {ex.Message}");
            CleanupTempFiles(exeDir);
            return false;
        }
    }

    /// <summary>
    /// Restart the application to apply the update.
    /// Preserves command line arguments (e.g., -headless).
    /// </summary>
    public static void RestartApp()
    {
        try
        {
            var args = Environment.GetCommandLineArgs();
            var argsToPass = args.Length > 1
                ? string.Join(" ", args.Skip(1).Select(EscapeArg))
                : "";

            Logger.Trace($"Restarting app for update (args: '{argsToPass}')...");

            // Remove MOTW from the exe we're about to launch
            RemoveZoneIdentifier(Application.ExecutablePath);

            var startInfo = new ProcessStartInfo
            {
                FileName = Application.ExecutablePath,
                Arguments = argsToPass,
                UseShellExecute = true
            };
            Process.Start(startInfo);
            Application.Exit();
        }
        catch (Exception ex)
        {
            Logger.Trace($"Restart failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Clean up old exe from previous update and recover from interrupted updates.
    /// Call on startup. Returns true if cleanup happened (meaning we just updated).
    /// </summary>
    public static bool CleanupOldVersion()
    {
        try
        {
            var exePath = Application.ExecutablePath;
            var exeDir = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;
            var mainExePath = Path.Combine(exeDir, "MosaicTools.exe");
            var oldExePath = Path.Combine(exeDir, "MosaicTools_old.exe");
            var newExePath = Path.Combine(exeDir, "MosaicTools_new.exe");
            var zipPath = Path.Combine(exeDir, "MosaicTools_update.zip");
            var didCleanup = false;

            // Recovery: if main exe is missing, try to restore from _old or _new
            if (!File.Exists(mainExePath))
            {
                Logger.Trace("RECOVERY: Main exe missing after interrupted update");
                if (File.Exists(newExePath))
                {
                    try
                    {
                        MoveFileWithRetry(newExePath, mainExePath);
                        Logger.Trace("RECOVERY: Restored from _new.exe");
                    }
                    catch (Exception ex) { Logger.Trace($"RECOVERY from _new failed: {ex.Message}"); }
                }
                else if (File.Exists(oldExePath))
                {
                    try
                    {
                        MoveFileWithRetry(oldExePath, mainExePath);
                        Logger.Trace("RECOVERY: Restored from _old.exe");
                    }
                    catch (Exception ex) { Logger.Trace($"RECOVERY from _old failed: {ex.Message}"); }
                }
            }

            // Normal cleanup: _old.exe means we just updated successfully
            if (File.Exists(oldExePath))
            {
                // Write health marker — confirms the new version started OK
                try
                {
                    var healthPath = Path.Combine(exeDir, "MosaicTools_healthy.flag");
                    File.WriteAllText(healthPath, GetCurrentVersion().ToString());
                }
                catch { }

                DeleteFileWithRetry(oldExePath);
                Logger.Trace("Cleaned up old version - just updated");
                didCleanup = true;
            }

            // Clean up leftover _new.exe from interrupted update
            if (File.Exists(newExePath))
                DeleteFileWithRetry(newExePath);

            // Clean up leftover zip from interrupted update
            if (File.Exists(zipPath))
                DeleteFileWithRetry(zipPath);

            if (didCleanup) return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Cleanup failed: {ex.Message}");
        }
        return false;
    }

    /// <summary>
    /// Normalize version string to be parseable (e.g., "2.1" -> "2.1.0")
    /// </summary>
    private static string NormalizeVersion(string version)
    {
        var parts = version.Split('.');
        return parts.Length switch
        {
            1 => $"{parts[0]}.0.0",
            2 => $"{parts[0]}.{parts[1]}.0",
            _ => version
        };
    }

    /// <summary>
    /// Properly escape a command line argument for Process.Start.
    /// Handles embedded quotes and backslashes per Windows CommandLineToArgvW rules.
    /// </summary>
    private static string EscapeArg(string arg)
    {
        if (arg.Length > 0 && !arg.Any(c => c == ' ' || c == '"' || c == '\\'))
            return arg;

        // Wrap in quotes, escaping internal quotes and trailing backslashes
        var escaped = arg.Replace("\\\"", "\\\\\"").Replace("\"", "\\\"");
        if (escaped.EndsWith('\\'))
            escaped += "\\"; // double trailing backslash before closing quote
        return $"\"{escaped}\"";
    }

    private static bool IsTransientHttpError(System.Net.HttpStatusCode code)
    {
        return code == System.Net.HttpStatusCode.RequestTimeout
            || code == System.Net.HttpStatusCode.TooManyRequests
            || code == System.Net.HttpStatusCode.ServiceUnavailable
            || code == System.Net.HttpStatusCode.GatewayTimeout
            || (int)code >= 500;
    }

    private static bool IsTransientException(Exception ex)
    {
        return ex is HttpRequestException || ex is TaskCanceledException || ex is TimeoutException;
    }

    private static void CleanupTempFiles(string exeDir)
    {
        DeleteFileWithRetry(Path.Combine(exeDir, "MosaicTools_new.exe"));
        DeleteFileWithRetry(Path.Combine(exeDir, "MosaicTools_update.zip"));
    }
}

// GitHub API response models
public class GitHubRelease
{
    [JsonPropertyName("tag_name")]
    public string? TagName { get; set; }

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("body")]
    public string? Body { get; set; }

    [JsonPropertyName("assets")]
    public List<GitHubAsset>? Assets { get; set; }
}

public class GitHubAsset
{
    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("browser_download_url")]
    public string? BrowserDownloadUrl { get; set; }
}

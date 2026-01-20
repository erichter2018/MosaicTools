using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Json;
using System.Reflection;
using System.Text.Json.Serialization;
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
    private static readonly HttpClient _httpClient;

    public string? LatestVersion { get; private set; }
    public string? DownloadUrl { get; private set; }
    public string? ReleaseNotes { get; private set; }
    public bool UpdateReady { get; private set; }

    static UpdateService()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.Add("User-Agent", "MosaicTools-AutoUpdater");
        // Note: Timeout is set per-request in CheckForUpdateAsync and DownloadUpdateAsync
    }

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
    /// </summary>
    public async Task<bool> CheckForUpdateAsync()
    {
        try
        {
            Logger.Trace("Checking for updates...");

            // Use 15 second timeout for the API check
            using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(15));
            var response = await _httpClient.GetAsync(GitHubApiUrl, cts.Token);
            if (!response.IsSuccessStatusCode)
            {
                Logger.Trace($"GitHub API returned {response.StatusCode}");
                return false;
            }

            var release = await response.Content.ReadFromJsonAsync<GitHubRelease>();
            if (release == null)
            {
                Logger.Trace("Failed to parse GitHub release");
                return false;
            }

            // Parse version from tag (e.g., "v2.1" or "2.1.0")
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
                // Find the zip asset (preferred) or exe asset (fallback for old releases)
                var zipAsset = release.Assets?.Find(a =>
                    a.Name?.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) == true);

                var exeAsset = release.Assets?.Find(a =>
                    a.Name?.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) == true);

                var asset = zipAsset ?? exeAsset;

                if (asset != null)
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
        catch (Exception ex)
        {
            Logger.Trace($"Update check failed: {ex.Message}");
            return false;
        }
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

        try
        {
            var exePath = Application.ExecutablePath;
            var exeDir = Path.GetDirectoryName(exePath) ?? ".";
            var newExePath = Path.Combine(exeDir, "MosaicTools_new.exe");
            var oldExePath = Path.Combine(exeDir, "MosaicTools_old.exe");
            var isZipDownload = DownloadUrl.EndsWith(".zip", StringComparison.OrdinalIgnoreCase);
            var downloadPath = isZipDownload
                ? Path.Combine(exeDir, "MosaicTools_update.zip")
                : newExePath;

            Logger.Trace($"Downloading update from {DownloadUrl} (zip={isZipDownload})");

            // Download file (5 minute timeout for large files)
            using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromMinutes(5));
            using var response = await _httpClient.GetAsync(DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cts.Token);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength ?? -1;
            var downloadedBytes = 0L;

            await using var contentStream = await response.Content.ReadAsStreamAsync();
            await using var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None);

            var buffer = new byte[81920];
            int bytesRead;

            while ((bytesRead = await contentStream.ReadAsync(buffer)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead));
                downloadedBytes += bytesRead;

                if (totalBytes > 0)
                {
                    var progress = (int)(downloadedBytes * 100 / totalBytes);
                    progressCallback?.Invoke(progress);
                }
            }

            fileStream.Close();
            Logger.Trace($"Download complete: {downloadedBytes} bytes");

            // If it's a zip file, extract the exe from it
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
                        File.Delete(downloadPath);
                        return false;
                    }

                    // Extract to _new.exe
                    exeEntry.ExtractToFile(newExePath, overwrite: true);
                    Logger.Trace($"Extracted {exeEntry.Name} ({exeEntry.Length} bytes)");
                }
                finally
                {
                    // Clean up the zip file
                    try { File.Delete(downloadPath); }
                    catch { /* ignore */ }
                }
            }

            // Verify the download is a valid exe (basic check)
            var fileInfo = new FileInfo(newExePath);
            if (fileInfo.Length < 100000) // Expect at least 100KB
            {
                Logger.Trace("Downloaded file too small, likely invalid");
                File.Delete(newExePath);
                return false;
            }

            // Clean up any previous old exe
            if (File.Exists(oldExePath))
            {
                try { File.Delete(oldExePath); }
                catch { /* ignore */ }
            }

            // Rename current exe to _old (Windows allows renaming running exe)
            File.Move(exePath, oldExePath);

            // Rename new exe to current name
            File.Move(newExePath, exePath);

            Logger.Trace("Update files ready, restart required");
            UpdateReady = true;
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Download failed: {ex.Message}");

            // Try to clean up
            var exeDir = Path.GetDirectoryName(Application.ExecutablePath) ?? ".";
            var newExePath = Path.Combine(exeDir, "MosaicTools_new.exe");
            var zipPath = Path.Combine(exeDir, "MosaicTools_update.zip");

            if (File.Exists(newExePath))
            {
                try { File.Delete(newExePath); }
                catch { /* ignore */ }
            }
            if (File.Exists(zipPath))
            {
                try { File.Delete(zipPath); }
                catch { /* ignore */ }
            }

            return false;
        }
    }

    /// <summary>
    /// Restart the application to apply the update.
    /// </summary>
    public static void RestartApp()
    {
        try
        {
            Logger.Trace("Restarting app for update...");
            Process.Start(Application.ExecutablePath);
            Application.Exit();
        }
        catch (Exception ex)
        {
            Logger.Trace($"Restart failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Clean up old exe from previous update (call on startup).
    /// </summary>
    public static void CleanupOldVersion()
    {
        try
        {
            var exeDir = Path.GetDirectoryName(Application.ExecutablePath) ?? ".";
            var oldExePath = Path.Combine(exeDir, "MosaicTools_old.exe");

            if (File.Exists(oldExePath))
            {
                File.Delete(oldExePath);
                Logger.Trace("Cleaned up old version");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Cleanup failed: {ex.Message}");
        }
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

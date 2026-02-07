using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace MosaicTools.UI;

/// <summary>
/// Hosts the Mic Gain Calibrator HTML tool in a WebView2 control.
/// </summary>
public class AudioSetupForm : Form
{
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    private WebView2? _webView;

    public AudioSetupForm()
    {
        Text = "Mic Gain Calibrator";
        Size = new Size(850, 800);
        MinimumSize = new Size(600, 400);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(30, 30, 30);
        Icon = FindAppIcon();

        InitializeWebView();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        // Enable dark title bar
        int value = 1;
        DwmSetWindowAttribute(Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
    }

    private void InitializeWebView()
    {
        _webView = new WebView2
        {
            Dock = DockStyle.Fill,
            DefaultBackgroundColor = Color.FromArgb(8, 9, 12) // Match HTML --bg color
        };

        _webView.CoreWebView2InitializationCompleted += OnWebViewReady;
        Controls.Add(_webView);

        // Initialize WebView2 with a user data folder in AppData
        var userDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MosaicTools", "WebView2");
        var env = CoreWebView2Environment.CreateAsync(null, userDataFolder).GetAwaiter().GetResult();
        _webView.EnsureCoreWebView2Async(env);
    }

    private void OnWebViewReady(object? sender, CoreWebView2InitializationCompletedEventArgs e)
    {
        if (!e.IsSuccess || _webView?.CoreWebView2 == null)
        {
            Services.Logger.Trace($"WebView2 init failed: {e.InitializationException?.Message}");
            MessageBox.Show(
                "Could not initialize the browser component.\n\n" +
                "The WebView2 Runtime may not be installed.\n" +
                "Download it from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
                "Mic Gain Calibrator",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            Close();
            return;
        }

        // Auto-allow microphone access
        _webView.CoreWebView2.PermissionRequested += (s, args) =>
        {
            if (args.PermissionKind == CoreWebView2PermissionKind.Microphone)
                args.State = CoreWebView2PermissionState.Allow;
        };

        // Write HTML to temp folder and serve via virtual HTTPS host
        // (getUserMedia requires a secure origin, NavigateToString uses about:blank which blocks it)
        var html = LoadHtmlResource();
        if (html != null)
        {
            var contentDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MosaicTools", "WebView2", "content");
            Directory.CreateDirectory(contentDir);
            File.WriteAllText(Path.Combine(contentDir, "MicCalibrator.html"), html);

            _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "mic-calibrator.local", contentDir,
                CoreWebView2HostResourceAccessKind.Allow);
            _webView.CoreWebView2.Navigate("https://mic-calibrator.local/MicCalibrator.html");
        }
        else
        {
            _webView.CoreWebView2.NavigateToString(
                "<html><body style='background:#1e1e1e;color:white;font-family:sans-serif;padding:2rem;'>" +
                "<h2>Error</h2><p>Could not load MicCalibrator.html resource.</p></body></html>");
        }
    }

    private static string? LoadHtmlResource()
    {
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "MosaicTools.MicCalibrator.html";

            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
            {
                Services.Logger.Trace($"MicCalibrator resource not found: {resourceName}");
                return null;
            }

            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch (Exception ex)
        {
            Services.Logger.Trace($"Error loading MicCalibrator resource: {ex.Message}");
            return null;
        }
    }

    private static Icon? FindAppIcon()
    {
        try
        {
            var exePath = Application.ExecutablePath;
            return Icon.ExtractAssociatedIcon(exePath);
        }
        catch
        {
            return null;
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _webView?.Dispose();
        }
        base.Dispose(disposing);
    }
}

using System.Diagnostics;
using System.Runtime.Versioning;

namespace SaleCast.Printer.Services;

/// <summary>
/// macOS menu bar service - placeholder for future implementation
/// Note: Full menu bar support on macOS requires native Cocoa APIs or Avalonia
/// For now, this service just logs the API URL for the user
/// </summary>
[SupportedOSPlatform("macos")]
public class MacMenuBarService : IDisposable
{
    private readonly ILogger<MacMenuBarService> _logger;
    private readonly IHostApplicationLifetime _appLifetime;
    private readonly int _port;
    private bool _disposed;

    public MacMenuBarService(
        ILogger<MacMenuBarService> logger,
        IHostApplicationLifetime appLifetime,
        IConfiguration configuration)
    {
        _logger = logger;
        _appLifetime = appLifetime;
        _port = configuration.GetValue("Api:Port", 5123);

        _logger.LogInformation("SaleCast Printer Service running on macOS");
        _logger.LogInformation("API available at: http://localhost:{Port}", _port);
        _logger.LogInformation("Swagger UI: http://localhost:{Port}/swagger", _port);

        // On macOS, we could potentially use terminal-notifier or osascript for notifications
        // For now, we just run as a background service
    }

    /// <summary>
    /// Show a macOS notification using osascript (optional)
    /// </summary>
    public void ShowNotification(string title, string message)
    {
        try
        {
            var script = $"display notification \"{message}\" with title \"{title}\"";
            Process.Start(new ProcessStartInfo
            {
                FileName = "osascript",
                Arguments = $"-e '{script}'",
                UseShellExecute = false,
                CreateNoWindow = true
            });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to show macOS notification");
        }
    }

    /// <summary>
    /// Open Swagger UI in default browser
    /// </summary>
    public void OpenSwagger()
    {
        var url = $"http://localhost:{_port}/swagger";
        Process.Start(new ProcessStartInfo
        {
            FileName = "open",
            Arguments = url,
            UseShellExecute = false
        });
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        GC.SuppressFinalize(this);
    }
}

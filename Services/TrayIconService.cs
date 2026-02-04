using System.Diagnostics;
using System.Drawing;
using System.Runtime.Versioning;
using H.NotifyIcon;
using H.NotifyIcon.Core;

namespace SaleCast.Printer.Services;

[SupportedOSPlatform("windows")]
public class TrayIconService : IDisposable
{
    private readonly TrayIconWithContextMenu _trayIcon;
    private readonly ILogger<TrayIconService> _logger;
    private readonly IHostApplicationLifetime _appLifetime;
    private readonly int _port;
    private bool _disposed;

    public TrayIconService(
        ILogger<TrayIconService> logger,
        IHostApplicationLifetime appLifetime,
        IConfiguration configuration)
    {
        _logger = logger;
        _appLifetime = appLifetime;
        _port = configuration.GetValue("Api:Port", 5123);

        _trayIcon = new TrayIconWithContextMenu
        {
            ToolTip = "SaleCast Printer Service",
            Icon = LoadIcon().Handle
        };

        _trayIcon.ContextMenu = CreateContextMenu();
        _trayIcon.MessageWindow.MouseEventReceived += (_, e) =>
        {
            if (e.MouseEvent == MouseEvent.IconLeftMouseUp)
            {
                OpenSwagger();
            }
        };

        _trayIcon.Create();

        _logger.LogInformation("System tray icon initialized");
    }

    private static Icon LoadIcon()
    {
        // Use IconService to load icon (supports SVG, ICO, PNG)
        return IconService.LoadAppIcon(32);
    }

    private PopupMenu CreateContextMenu()
    {
        var menu = new PopupMenu();

        // Status header (disabled)
        menu.Items.Add(new PopupMenuItem($"SaleCast Printer v{GetVersion()}", (_, _) => { }));

        menu.Items.Add(new PopupMenuSeparator());

        // Open Swagger
        menu.Items.Add(new PopupMenuItem("Open API Documentation", (_, _) => OpenSwagger()));

        // Open Logs folder
        menu.Items.Add(new PopupMenuItem("Open Logs Folder", (_, _) => OpenLogsFolder()));

        menu.Items.Add(new PopupMenuSeparator());

        // List printers
        menu.Items.Add(new PopupMenuItem("View Printers (JSON)", (_, _) => OpenPrintersList()));

        menu.Items.Add(new PopupMenuSeparator());

        // Quit
        menu.Items.Add(new PopupMenuItem("Quit", (_, _) => Quit()));

        return menu;
    }

    private void OpenSwagger()
    {
        var url = $"http://localhost:{_port}/swagger";
        _logger.LogInformation("Opening Swagger UI: {Url}", url);

        Process.Start(new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
    }

    private void OpenPrintersList()
    {
        var url = $"http://localhost:{_port}/printers";
        _logger.LogInformation("Opening printers list: {Url}", url);

        Process.Start(new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
    }

    private void OpenLogsFolder()
    {
        var logsPath = Path.Combine(AppContext.BaseDirectory, "logs");
        Directory.CreateDirectory(logsPath);

        _logger.LogInformation("Opening logs folder: {Path}", logsPath);

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = logsPath,
            UseShellExecute = true
        });
    }

    private void Quit()
    {
        _logger.LogInformation("Quit requested from tray menu");
        _appLifetime.StopApplication();
    }

    private static string GetVersion()
    {
        return typeof(TrayIconService).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";
    }

    public void Dispose()
    {
        if (_disposed) return;

        _trayIcon.Dispose();
        _disposed = true;

        GC.SuppressFinalize(this);
    }
}

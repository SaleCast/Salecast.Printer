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
        // Try to load custom icon from Assets folder
        var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "printer.ico");
        if (File.Exists(iconPath))
        {
            return new Icon(iconPath);
        }

        // Try to extract icon from the executable
        var exePath = Environment.ProcessPath;
        if (exePath != null && File.Exists(exePath))
        {
            var icon = Icon.ExtractAssociatedIcon(exePath);
            if (icon != null) return icon;
        }

        // Fallback: create a simple printer icon programmatically
        return CreatePrinterIcon();
    }

    private static Icon CreatePrinterIcon()
    {
        using var bitmap = new Bitmap(32, 32);
        using var g = Graphics.FromImage(bitmap);

        // Background
        g.Clear(Color.Transparent);

        // Printer body (gray box)
        using var bodyBrush = new SolidBrush(Color.FromArgb(80, 80, 80));
        g.FillRectangle(bodyBrush, 4, 10, 24, 14);

        // Paper tray (top)
        using var paperBrush = new SolidBrush(Color.White);
        g.FillRectangle(paperBrush, 8, 4, 16, 8);

        // Output paper (bottom)
        g.FillRectangle(paperBrush, 8, 22, 16, 6);

        // Printer details
        using var detailBrush = new SolidBrush(Color.FromArgb(60, 60, 60));
        g.FillRectangle(detailBrush, 6, 14, 4, 4);

        // Green status light
        using var statusBrush = new SolidBrush(Color.LimeGreen);
        g.FillEllipse(statusBrush, 20, 14, 4, 4);

        return Icon.FromHandle(bitmap.GetHicon());
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

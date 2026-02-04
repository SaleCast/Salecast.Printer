using System.Reflection;
using System.Runtime.Versioning;
using Microsoft.Win32;

namespace SaleCast.Printer.Services;

[SupportedOSPlatform("windows")]
public class WindowsStartupService : IStartupService
{
    private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string AppName = "SaleCast.Printer";

    private readonly ILogger<WindowsStartupService> _logger;
    private readonly IConfiguration _configuration;

    public WindowsStartupService(ILogger<WindowsStartupService> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
    }

    public bool IsRegisteredForStartup()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false);
            var value = key?.GetValue(AppName);
            return value != null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to check startup registration");
            return false;
        }
    }

    public void RegisterForStartup()
    {
        try
        {
            var exePath = Environment.ProcessPath ?? Assembly.GetExecutingAssembly().Location;
            var command = $"\"{exePath}\" --minimized";

            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);
            key?.SetValue(AppName, command);

            _logger.LogInformation("Registered for startup: {Command}", command);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to register for startup");
        }
    }

    public void UnregisterFromStartup()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);
            key?.DeleteValue(AppName, false);

            _logger.LogInformation("Unregistered from startup");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to unregister from startup");
        }
    }

    public void EnsureStartupRegistration()
    {
        var runAtBoot = _configuration.GetValue("Startup:RunAtBoot", true);

        if (runAtBoot && !IsRegisteredForStartup())
        {
            RegisterForStartup();
        }
        else if (!runAtBoot && IsRegisteredForStartup())
        {
            UnregisterFromStartup();
        }
    }
}

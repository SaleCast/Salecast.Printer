using System.Reflection;
using System.Xml.Linq;

namespace SaleCast.Printer.Services;

/// <summary>
/// macOS startup service using LaunchAgents
/// </summary>
public class MacStartupService : IStartupService
{
    private const string AppIdentifier = "com.salecast.printer";
    private readonly ILogger<MacStartupService> _logger;
    private readonly IConfiguration _configuration;

    public MacStartupService(ILogger<MacStartupService> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
    }

    private static string LaunchAgentPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Library", "LaunchAgents", $"{AppIdentifier}.plist"
        );

    public bool IsRegisteredForStartup()
    {
        return File.Exists(LaunchAgentPath);
    }

    public void RegisterForStartup()
    {
        try
        {
            var exePath = Environment.ProcessPath ?? Assembly.GetExecutingAssembly().Location;
            var launchAgentDir = Path.GetDirectoryName(LaunchAgentPath)!;

            Directory.CreateDirectory(launchAgentDir);

            // Create plist content
            var plist = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XDocumentType("plist", "-//Apple//DTD PLIST 1.0//EN", "http://www.apple.com/DTDs/PropertyList-1.0.dtd", null),
                new XElement("plist",
                    new XAttribute("version", "1.0"),
                    new XElement("dict",
                        new XElement("key", "Label"),
                        new XElement("string", AppIdentifier),
                        new XElement("key", "ProgramArguments"),
                        new XElement("array",
                            new XElement("string", exePath),
                            new XElement("string", "--minimized")
                        ),
                        new XElement("key", "RunAtLoad"),
                        new XElement("true"),
                        new XElement("key", "KeepAlive"),
                        new XElement("false"),
                        new XElement("key", "StandardOutPath"),
                        new XElement("string", Path.Combine(launchAgentDir, "salecast-printer.log")),
                        new XElement("key", "StandardErrorPath"),
                        new XElement("string", Path.Combine(launchAgentDir, "salecast-printer-error.log"))
                    )
                )
            );

            plist.Save(LaunchAgentPath);

            _logger.LogInformation("Registered for startup at {Path}", LaunchAgentPath);
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
            if (File.Exists(LaunchAgentPath))
            {
                File.Delete(LaunchAgentPath);
                _logger.LogInformation("Unregistered from startup");
            }
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

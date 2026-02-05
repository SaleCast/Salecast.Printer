using SaleCast.Printer.Models;
using SaleCast.Printer.Services;
using Serilog;
using Velopack;

// Velopack: Handle install/update hooks
VelopackApp.Build()
#if WINDOWS
    .OnFirstRun(_ => RegisterWindowsStartup())
#endif
    .Run();

#if WINDOWS
// Helper method for startup registration (called by Velopack hook)
static void RegisterWindowsStartup()
{
    var exePath = Environment.ProcessPath;
    if (exePath != null)
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Run", true);
            key?.SetValue("SaleCast.Printer", $"\"{exePath}\" --minimized");
        }
        catch { /* Will be handled by startup service */ }
    }
}
#endif

var builder = WebApplication.CreateBuilder(args);

// Configure Serilog
Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Configuration)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File(
        Path.Combine(AppContext.BaseDirectory, "logs", "salecast-printer-.log"),
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: 7)
    .CreateLogger();

builder.Host.UseSerilog();

// Configure port
var port = builder.Configuration.GetValue("Api:Port", 5123);
builder.WebHost.UseUrls($"http://localhost:{port}");

// Register platform-specific services
#if WINDOWS
builder.Services.AddSingleton<IPrinterService, WindowsPrinterService>();
builder.Services.AddSingleton<WindowsStartupService>();
builder.Services.AddSingleton<TrayIconService>();
#elif MACOS
builder.Services.AddSingleton<IPrinterService, MacPrinterService>();
builder.Services.AddSingleton<MacStartupService>();
builder.Services.AddSingleton<MacMenuBarService>();
#else
throw new PlatformNotSupportedException("Only Windows and macOS are supported");
#endif

// Add Swagger/OpenAPI
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
    {
        Title = "SaleCast Printer API",
        Version = "v1",
        Description = "Local printing service for SaleCast. Supports PDF and ZPL document formats."
    });
});

// Configure CORS
var allowedOrigins = builder.Configuration.GetSection("Api:AllowedOrigins").Get<string[]>() ?? [];
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins(allowedOrigins)
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

var app = builder.Build();

// Enable Swagger (always, since this is a local service)
app.UseSwagger();
app.UseSwaggerUI(options =>
{
    options.SwaggerEndpoint("/swagger/v1/swagger.json", "SaleCast Printer API v1");
    options.DocumentTitle = "SaleCast Printer API";
});

app.UseCors();
app.UseSerilogRequestLogging();

// Health check endpoint
app.MapGet("/", () => new { status = "ok", version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "1.0.0" })
    .WithName("HealthCheck")
    .WithDescription("Check if the service is running")
    .WithTags("Health");

// GET /printers - List all printers
app.MapGet("/printers", (IPrinterService printerService) =>
{
    var printers = printerService.GetPrinters();
    return Results.Ok(printers);
})
.WithName("GetPrinters")
.WithDescription("Get list of all available printers on this computer")
.WithTags("Printers")
.Produces<List<PrinterInfo>>();

// POST /printers/{printerId}/print - Print a document (PDF or ZPL)
app.MapPost("/printers/{printerId}/print", async (
    string printerId,
    IFormFile file,
    string documentType,
    string paperFormat,
    int? copies,
    IPrinterService printerService) =>
{
    // Validate printer exists
    if (!printerService.PrinterExists(printerId))
    {
        return Results.NotFound(new PrintResponse(false, null, $"Printer '{printerId}' not found"));
    }

    // Validate document type
    if (!Enum.TryParse<DocumentType>(documentType, ignoreCase: true, out var docType))
    {
        var validTypes = string.Join(", ", Enum.GetNames<DocumentType>());
        return Results.BadRequest(new PrintResponse(false, null, $"Invalid document type. Valid types: {validTypes}"));
    }

    // Validate paper format
    if (!Enum.TryParse<PaperFormat>(paperFormat, ignoreCase: true, out var format))
    {
        var validFormats = string.Join(", ", Enum.GetNames<PaperFormat>());
        return Results.BadRequest(new PrintResponse(false, null, $"Invalid paper format. Valid formats: {validFormats}"));
    }

    // Validate file
    if (file.Length == 0)
    {
        return Results.BadRequest(new PrintResponse(false, null, "File is empty"));
    }

    // Print document
    await using var stream = file.OpenReadStream();
    var response = printerService.PrintDocument(printerId, stream, docType, format, copies ?? 1);

    return response.Success ? Results.Ok(response) : Results.StatusCode(500);
})
.WithName("PrintDocument")
.WithDescription("Print a PDF or ZPL document to the specified printer")
.WithTags("Printers")
.Produces<PrintResponse>()
.Produces<PrintResponse>(StatusCodes.Status400BadRequest)
.Produces<PrintResponse>(StatusCodes.Status404NotFound)
.DisableAntiforgery();

// Run startup tasks
await using (var scope = app.Services.CreateAsyncScope())
{
    var logger = scope.ServiceProvider.GetRequiredService<ILogger<Program>>();

    // Register for OS startup if configured
#if WINDOWS
    var startupService = scope.ServiceProvider.GetRequiredService<WindowsStartupService>();
    startupService.EnsureStartupRegistration();

    // Initialize tray icon (keeps it alive)
    _ = scope.ServiceProvider.GetRequiredService<TrayIconService>();
#elif MACOS
    var startupService = scope.ServiceProvider.GetRequiredService<MacStartupService>();
    startupService.EnsureStartupRegistration();
#endif

    // Check for Velopack updates periodically (every 6 hours)
    var updateUrl = builder.Configuration["Update:Url"];
    if (!string.IsNullOrEmpty(updateUrl))
    {
        _ = Task.Run(async () =>
        {
            var updateCheckInterval = TimeSpan.FromHours(6);
            using var timer = new PeriodicTimer(updateCheckInterval);

            // Check immediately on startup, then every 6 hours
            do
            {
                await CheckForUpdatesAsync(updateUrl, logger);
            }
            while (await timer.WaitForNextTickAsync());
        });
    }
    else
    {
        logger.LogInformation("Update URL not configured, auto-update disabled");
    }
}

static async Task CheckForUpdatesAsync(string updateUrl, Microsoft.Extensions.Logging.ILogger logger)
{
    try
    {
        logger.LogInformation("Checking for updates...");
        var mgr = new UpdateManager(updateUrl);
        var updateInfo = await mgr.CheckForUpdatesAsync();

        if (updateInfo != null)
        {
            logger.LogInformation("Update available: {Version}", updateInfo.TargetFullRelease.Version);
            await mgr.DownloadUpdatesAsync(updateInfo);
            logger.LogInformation("Update downloaded, restarting to apply...");
            mgr.ApplyUpdatesAndRestart(updateInfo);
        }
        else
        {
            logger.LogInformation("Application is up to date");
        }
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "Error checking for updates");
    }
}

Log.Information("SaleCast.Printer started on port {Port}", port);
Log.Information("Swagger UI available at http://localhost:{Port}/swagger", port);

#if MACOS
// macOS: Start web host in background, keep main thread for AppKit event loop
await app.StartAsync();
try
{
    var menuBarService = app.Services.GetRequiredService<MacMenuBarService>();
    menuBarService.Initialize();
    menuBarService.Run(); // Blocks main thread until Quit
}
finally
{
    await app.StopAsync();
}
#else
app.Run();
#endif

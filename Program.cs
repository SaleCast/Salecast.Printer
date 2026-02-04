using SaleCast.Printer.Models;
using SaleCast.Printer.Services;
using Serilog;
using Velopack;

// Velopack: Handle install/update hooks
VelopackApp.Build()
    .OnFirstRun(v =>
    {
        // Register for startup on first install
        if (OperatingSystem.IsWindows())
        {
            RegisterWindowsStartup();
        }
    })
    .Run();

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
if (OperatingSystem.IsWindows())
{
    builder.Services.AddSingleton<IPrinterService, WindowsPrinterService>();
    builder.Services.AddSingleton<WindowsStartupService>();
    builder.Services.AddSingleton<TrayIconService>();
}
else if (OperatingSystem.IsMacOS())
{
    builder.Services.AddSingleton<IPrinterService, MacPrinterService>();
    builder.Services.AddSingleton<MacStartupService>();
}
else
{
    throw new PlatformNotSupportedException("Only Windows and macOS are supported");
}

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
    if (OperatingSystem.IsWindows())
    {
        var startupService = scope.ServiceProvider.GetRequiredService<WindowsStartupService>();
        startupService.EnsureStartupRegistration();

        // Initialize tray icon (keeps it alive)
        _ = scope.ServiceProvider.GetRequiredService<TrayIconService>();
    }
    else if (OperatingSystem.IsMacOS())
    {
        var startupService = scope.ServiceProvider.GetRequiredService<MacStartupService>();
        startupService.EnsureStartupRegistration();
    }

    // Check for Velopack updates (non-blocking)
    _ = Task.Run(async () =>
    {
        try
        {
            var updateUrl = builder.Configuration["Update:Url"];
            if (string.IsNullOrEmpty(updateUrl))
            {
                logger.LogInformation("Update URL not configured, skipping update check");
                return;
            }

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
    });
}

Log.Information("SaleCast.Printer started on port {Port}", port);
Log.Information("Swagger UI available at http://localhost:{Port}/swagger", port);

app.Run();

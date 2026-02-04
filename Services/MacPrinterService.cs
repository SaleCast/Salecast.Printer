using System.Diagnostics;
using System.Text.RegularExpressions;
using SaleCast.Printer.Models;

namespace SaleCast.Printer.Services;

/// <summary>
/// macOS printer service using CUPS (Common Unix Printing System)
/// </summary>
public partial class MacPrinterService : IPrinterService
{
    private readonly ILogger<MacPrinterService> _logger;

    public MacPrinterService(ILogger<MacPrinterService> logger)
    {
        _logger = logger;
    }

    public List<PrinterInfo> GetPrinters()
    {
        var printers = new List<PrinterInfo>();

        try
        {
            var output = RunCommand("lpstat", "-p -d");
            var defaultPrinter = GetDefaultPrinterName(output);

            var printerRegex = PrinterNameRegex();
            var matches = printerRegex.Matches(output);

            foreach (Match match in matches)
            {
                var name = match.Groups[1].Value;
                printers.Add(new PrinterInfo(
                    Id: name,
                    Name: name,
                    Default: name.Equals(defaultPrinter, StringComparison.OrdinalIgnoreCase)
                ));
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to enumerate printers");
        }

        return printers;
    }

    public bool PrinterExists(string printerId)
    {
        var printers = GetPrinters();
        return printers.Any(p => p.Id.Equals(printerId, StringComparison.OrdinalIgnoreCase));
    }

    public PrintResponse PrintDocument(string printerId, Stream fileStream, DocumentType documentType, PaperFormat format, int copies = 1)
    {
        return documentType switch
        {
            DocumentType.ZPL => PrintZpl(printerId, fileStream, copies),
            DocumentType.PDF => PrintPdf(printerId, fileStream, format, copies),
            _ => new PrintResponse(false, null, $"Unsupported document type: {documentType}")
        };
    }

    private PrintResponse PrintZpl(string printerId, Stream fileStream, int copies)
    {
        try
        {
            var tempFile = Path.Combine(Path.GetTempPath(), $"salecast_zpl_{Guid.NewGuid():N}.zpl");

            using (var file = File.Create(tempFile))
            {
                fileStream.CopyTo(file);
            }

            try
            {
                // Use lp with raw option for ZPL
                var args = $"-d \"{printerId}\" -n {copies} -o raw \"{tempFile}\"";
                var output = RunCommand("lp", args);

                var jobIdMatch = Regex.Match(output, @"request id is (\S+-\d+)");
                var jobId = jobIdMatch.Success ? jobIdMatch.Groups[1].Value : Guid.NewGuid().ToString("N")[..8];

                _logger.LogInformation("ZPL printed successfully to {Printer}, job {JobId} ({Copies} copies)",
                    printerId, jobId, copies);

                return new PrintResponse(
                    Success: true,
                    JobId: jobId,
                    Message: $"ZPL sent to printer ({copies} copies)"
                );
            }
            finally
            {
                try { File.Delete(tempFile); } catch { }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to print ZPL to {Printer}", printerId);
            return new PrintResponse(
                Success: false,
                JobId: null,
                Message: $"ZPL print failed: {ex.Message}"
            );
        }
    }

    private PrintResponse PrintPdf(string printerId, Stream fileStream, PaperFormat format, int copies)
    {
        try
        {
            var tempFile = Path.Combine(Path.GetTempPath(), $"salecast_print_{Guid.NewGuid():N}.pdf");

            using (var file = File.Create(tempFile))
            {
                fileStream.CopyTo(file);
            }

            try
            {
                var mediaSize = GetCupsMediaSize(format);
                var args = $"-d \"{printerId}\" -n {copies} -o media={mediaSize} \"{tempFile}\"";

                var output = RunCommand("lp", args);

                var jobIdMatch = Regex.Match(output, @"request id is (\S+-\d+)");
                var jobId = jobIdMatch.Success ? jobIdMatch.Groups[1].Value : Guid.NewGuid().ToString("N")[..8];

                _logger.LogInformation("PDF printed successfully to {Printer} with format {Format}, job {JobId}",
                    printerId, format, jobId);

                return new PrintResponse(
                    Success: true,
                    JobId: jobId,
                    Message: "Document sent to printer"
                );
            }
            finally
            {
                try { File.Delete(tempFile); } catch { }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to print PDF to {Printer}", printerId);
            return new PrintResponse(
                Success: false,
                JobId: null,
                Message: $"Print failed: {ex.Message}"
            );
        }
    }

    private static string? GetDefaultPrinterName(string lpstatOutput)
    {
        var match = Regex.Match(lpstatOutput, @"system default destination:\s*(\S+)");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string GetCupsMediaSize(PaperFormat format)
    {
        return format switch
        {
            PaperFormat.A4 => "A4",
            PaperFormat.A5 => "A5",
            PaperFormat.A6 => "A6",
            PaperFormat.A7 => "A7",
            PaperFormat.B5 => "ISOB5",
            PaperFormat.B6 => "ISOB6",
            PaperFormat.LEGAL => "Legal",
            PaperFormat.LETTER => "Letter",
            PaperFormat.LBL_4X6 => "4x6",
            PaperFormat.LBL_4X8 => "4x8",
            PaperFormat.LBL_4X4 => "4x4",
            PaperFormat.DHL_910_300_600 => "Custom.3x6in",
            _ => "Letter"
        };
    }

    private string RunCommand(string command, string arguments)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = command,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            }
        };

        process.Start();
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0 && !string.IsNullOrEmpty(error))
        {
            _logger.LogWarning("Command {Command} returned error: {Error}", command, error);
        }

        return output + error;
    }

    [GeneratedRegex(@"printer (\S+)", RegexOptions.Multiline)]
    private static partial Regex PrinterNameRegex();
}

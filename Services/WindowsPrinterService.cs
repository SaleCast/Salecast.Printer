using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Docnet.Core;
using Docnet.Core.Models;
using SaleCast.Printer.Models;

namespace SaleCast.Printer.Services;

[SupportedOSPlatform("windows")]
public class WindowsPrinterService : IPrinterService
{
    private readonly ILogger<WindowsPrinterService> _logger;

    public WindowsPrinterService(ILogger<WindowsPrinterService> logger)
    {
        _logger = logger;
    }

    public List<PrinterInfo> GetPrinters()
    {
        var printers = new List<PrinterInfo>();
        var defaultPrinter = GetDefaultPrinterName();

        foreach (string printerName in PrinterSettings.InstalledPrinters)
        {
            printers.Add(new PrinterInfo(
                Id: printerName,
                Name: printerName,
                Default: printerName.Equals(defaultPrinter, StringComparison.OrdinalIgnoreCase)
            ));
        }

        return printers;
    }

    public bool PrinterExists(string printerId)
    {
        foreach (string printerName in PrinterSettings.InstalledPrinters)
        {
            if (printerName.Equals(printerId, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
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
            using var memoryStream = new MemoryStream();
            fileStream.CopyTo(memoryStream);
            var zplData = memoryStream.ToArray();

            for (var i = 0; i < copies; i++)
            {
                if (!RawPrinterHelper.SendBytesToPrinter(printerId, zplData))
                {
                    return new PrintResponse(
                        Success: false,
                        JobId: null,
                        Message: "Failed to send ZPL data to printer"
                    );
                }
            }

            _logger.LogInformation("ZPL document printed successfully to {Printer} ({Copies} copies)", printerId, copies);

            return new PrintResponse(
                Success: true,
                JobId: Guid.NewGuid().ToString("N")[..8],
                Message: $"ZPL sent to printer ({copies} copies)"
            );
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
            using var memoryStream = new MemoryStream();
            fileStream.CopyTo(memoryStream);
            var pdfBytes = memoryStream.ToArray();

            var (paperWidth, paperHeight) = GetPaperDimensions(format);
            var renderWidth = (int)(paperWidth * 1.5); // 150 DPI
            var renderHeight = (int)(paperHeight * 1.5);

            using var docReader = DocLib.Instance.GetDocReader(pdfBytes, new PageDimensions(renderWidth, renderHeight));
            var pageCount = docReader.GetPageCount();

            using var printDocument = new PrintDocument();
            printDocument.PrinterSettings.PrinterName = printerId;
            printDocument.PrinterSettings.Copies = (short)Math.Min(copies, short.MaxValue);

            var paperSize = GetPaperSize(format, printDocument.PrinterSettings);
            if (paperSize != null)
            {
                printDocument.DefaultPageSettings.PaperSize = paperSize;
            }

            var currentPage = 0;

            printDocument.PrintPage += (sender, e) =>
            {
                using var pageReader = docReader.GetPageReader(currentPage);
                var rawBytes = pageReader.GetImage();
                var width = pageReader.GetPageWidth();
                var height = pageReader.GetPageHeight();

                using var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                var bitmapData = bitmap.LockBits(
                    new Rectangle(0, 0, width, height),
                    ImageLockMode.WriteOnly,
                    PixelFormat.Format32bppArgb);

                Marshal.Copy(rawBytes, 0, bitmapData.Scan0, rawBytes.Length);
                bitmap.UnlockBits(bitmapData);

                var pageWidth = e.PageBounds.Width;
                var pageHeight = e.PageBounds.Height;
                var scale = Math.Min((float)pageWidth / width, (float)pageHeight / height);
                var scaledWidth = (int)(width * scale);
                var scaledHeight = (int)(height * scale);
                var x = (pageWidth - scaledWidth) / 2;
                var y = (pageHeight - scaledHeight) / 2;

                e.Graphics!.DrawImage(bitmap, x, y, scaledWidth, scaledHeight);

                currentPage++;
                e.HasMorePages = currentPage < pageCount;
            };

            printDocument.PrintController = new StandardPrintController();
            printDocument.Print();

            _logger.LogInformation("PDF ({PageCount} pages) printed successfully to {Printer} with format {Format}",
                pageCount, printerId, format);

            return new PrintResponse(
                Success: true,
                JobId: Guid.NewGuid().ToString("N")[..8],
                Message: $"Document sent to printer ({pageCount} pages)"
            );
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

    private static string? GetDefaultPrinterName()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer WHERE Default = TRUE");
            foreach (var printer in searcher.Get())
            {
                return printer["Name"]?.ToString();
            }
        }
        catch
        {
            // Fallback if WMI query fails
        }
        return null;
    }

    private static PaperSize? GetPaperSize(PaperFormat format, PrinterSettings printerSettings)
    {
        var (width, height) = GetPaperDimensions(format);

        foreach (PaperSize paperSize in printerSettings.PaperSizes)
        {
            if (Math.Abs(paperSize.Width - width) < 10 && Math.Abs(paperSize.Height - height) < 10)
            {
                return paperSize;
            }
        }

        return new PaperSize(format.ToString(), width, height);
    }

    private static (int Width, int Height) GetPaperDimensions(PaperFormat format)
    {
        return format switch
        {
            PaperFormat.A4 => (827, 1169),
            PaperFormat.A5 => (583, 827),
            PaperFormat.A6 => (413, 583),
            PaperFormat.A7 => (291, 413),
            PaperFormat.B5 => (693, 984),
            PaperFormat.B6 => (492, 693),
            PaperFormat.LEGAL => (850, 1400),
            PaperFormat.LETTER => (850, 1100),
            PaperFormat.LBL_4X6 => (400, 600),
            PaperFormat.LBL_4X8 => (400, 800),
            PaperFormat.LBL_4X4 => (400, 400),
            PaperFormat.DHL_910_300_600 => (300, 600),
            _ => (850, 1100)
        };
    }
}

/// <summary>
/// Windows raw printing helper using winspool.drv
/// </summary>
[SupportedOSPlatform("windows")]
internal static class RawPrinterHelper
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private class DOCINFOW
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pDocName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pOutputFile;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pDatatype;
    }

    [DllImport("winspool.drv", EntryPoint = "OpenPrinterW", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool OpenPrinter(string szPrinter, out nint hPrinter, nint pd);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool ClosePrinter(nint hPrinter);

    [DllImport("winspool.drv", EntryPoint = "StartDocPrinterW", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int StartDocPrinter(nint hPrinter, int level, [In] DOCINFOW di);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool EndDocPrinter(nint hPrinter);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool StartPagePrinter(nint hPrinter);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool EndPagePrinter(nint hPrinter);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool WritePrinter(nint hPrinter, nint pBytes, int dwCount, out int dwWritten);

    public static bool SendBytesToPrinter(string printerName, byte[] bytes)
    {
        var pUnmanagedBytes = Marshal.AllocCoTaskMem(bytes.Length);
        Marshal.Copy(bytes, 0, pUnmanagedBytes, bytes.Length);

        var di = new DOCINFOW
        {
            pDocName = "SaleCast ZPL Document",
            pDatatype = "RAW"
        };

        var success = false;

        if (OpenPrinter(printerName, out var hPrinter, nint.Zero))
        {
            if (StartDocPrinter(hPrinter, 1, di) != 0)
            {
                if (StartPagePrinter(hPrinter))
                {
                    success = WritePrinter(hPrinter, pUnmanagedBytes, bytes.Length, out _);
                    EndPagePrinter(hPrinter);
                }
                EndDocPrinter(hPrinter);
            }
            ClosePrinter(hPrinter);
        }

        Marshal.FreeCoTaskMem(pUnmanagedBytes);
        return success;
    }
}

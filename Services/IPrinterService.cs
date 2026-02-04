using SaleCast.Printer.Models;

namespace SaleCast.Printer.Services;

public interface IPrinterService
{
    List<PrinterInfo> GetPrinters();
    bool PrinterExists(string printerId);
    PrintResponse PrintDocument(string printerId, Stream fileStream, DocumentType documentType, PaperFormat format, int copies = 1);
}

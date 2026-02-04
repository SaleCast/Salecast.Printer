namespace SaleCast.Printer.Services;

public interface IStartupService
{
    bool IsRegisteredForStartup();
    void RegisterForStartup();
    void UnregisterFromStartup();
}

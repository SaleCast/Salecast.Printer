using System.Text.Json.Serialization;

namespace SaleCast.Printer.Models;

public record PrinterInfo(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("default")] bool Default
);

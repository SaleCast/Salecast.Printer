using System.Text.Json.Serialization;

namespace SaleCast.Printer.Models;

public record PrintResponse(
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("jobId")] string? JobId,
    [property: JsonPropertyName("message")] string Message
);

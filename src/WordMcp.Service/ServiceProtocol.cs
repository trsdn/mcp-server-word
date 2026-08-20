using System.Text.Json;
using System.Text.Json.Serialization;

namespace WordMcp.Service;

/// <summary>
/// Wire format shared by every client and the session service.
/// </summary>
/// <remarks>
/// <para>Messages are newline-delimited JSON: one <see cref="ServiceRequest"/> per line in,
/// one <see cref="ServiceResponse"/> per line out. That keeps the framing trivial enough to
/// read with a text tool while debugging and avoids a JSON-RPC dependency for a protocol
/// with a single method.</para>
/// <para>Because the payload is line delimited, a serialized message must never contain a raw
/// newline. <see cref="Serialize{T}"/> therefore writes compact JSON, in which a newline inside
/// a string value is escaped.</para>
/// </remarks>
public static class ServiceProtocol
{
    /// <summary>
    /// Gets the serializer options used for every message on the wire.
    /// </summary>
    public static JsonSerializerOptions JsonOptions { get; } = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    /// <summary>
    /// Serializes a message to a single JSON line.
    /// </summary>
    /// <typeparam name="T">Type of the message.</typeparam>
    /// <param name="message">The message to serialize.</param>
    /// <returns>Compact JSON without a trailing newline.</returns>
    public static string Serialize<T>(T message)
        => JsonSerializer.Serialize(message, JsonOptions);

    /// <summary>
    /// Deserializes a message from a JSON line.
    /// </summary>
    /// <typeparam name="T">Type of the message.</typeparam>
    /// <param name="json">The JSON text.</param>
    /// <returns>The deserialized message, or <c>null</c> when the line was blank.</returns>
    public static T? Deserialize<T>(string json)
        => string.IsNullOrWhiteSpace(json) ? default : JsonSerializer.Deserialize<T>(json, JsonOptions);
}

/// <summary>
/// A single command sent to the session service.
/// </summary>
public sealed class ServiceRequest
{
    /// <summary>
    /// Gets the command to run, for example <c>session.open</c> or <c>service.ping</c>.
    /// </summary>
    public required string Command { get; init; }

    /// <summary>
    /// Gets the session the command applies to, when it is session scoped.
    /// </summary>
    public string? SessionId { get; init; }

    /// <summary>
    /// Gets the JSON-serialized command arguments.
    /// </summary>
    public string? Args { get; init; }

    /// <summary>
    /// Gets a free-form label describing who sent the request. Used for diagnostics only.
    /// </summary>
    public string? Source { get; init; }
}

/// <summary>
/// The result of a single command.
/// </summary>
public sealed class ServiceResponse
{
    /// <summary>
    /// Gets a value indicating whether the command succeeded.
    /// </summary>
    public bool Success { get; init; }

    /// <summary>
    /// Gets the error description when <see cref="Success"/> is <c>false</c>.
    /// </summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Gets the CLR type name of the failure, so callers can react to specific errors.
    /// </summary>
    public string? ErrorType { get; init; }

    /// <summary>
    /// Gets the JSON-serialized result payload when <see cref="Success"/> is <c>true</c>.
    /// </summary>
    public string? Result { get; init; }

    /// <summary>
    /// Creates a successful response carrying a serialized payload.
    /// </summary>
    /// <param name="payload">The value to return to the caller.</param>
    /// <returns>A successful response.</returns>
    public static ServiceResponse Ok(object payload)
        => new() { Success = true, Result = ServiceProtocol.Serialize(payload) };

    /// <summary>
    /// Creates a failed response.
    /// </summary>
    /// <param name="message">Description of the failure.</param>
    /// <param name="errorType">CLR type name of the failure.</param>
    /// <returns>A failed response.</returns>
    public static ServiceResponse Fail(string message, string? errorType = null)
        => new() { Success = false, ErrorMessage = message, ErrorType = errorType };
}

/// <summary>
/// A snapshot of what the service is currently doing.
/// </summary>
public sealed class ServiceStatus
{
    /// <summary>
    /// Gets the process id of the service.
    /// </summary>
    public required int ProcessId { get; init; }

    /// <summary>
    /// Gets the number of currently open sessions.
    /// </summary>
    public required int SessionCount { get; init; }

    /// <summary>
    /// Gets the UTC time the service started.
    /// </summary>
    public required DateTimeOffset StartedAt { get; init; }

    /// <summary>
    /// Gets the UTC time the service last handled a request.
    /// </summary>
    public required DateTimeOffset LastActivityAt { get; init; }

    /// <summary>
    /// Gets the idle period after which the service shuts itself down.
    /// </summary>
    public required TimeSpan IdleTimeout { get; init; }

    /// <summary>
    /// Gets the service version.
    /// </summary>
    public required string Version { get; init; }
}

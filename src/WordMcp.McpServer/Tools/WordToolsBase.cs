using System.Text.Json;
using System.Text.Json.Serialization;
using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;

namespace WordMcp.McpServer.Tools;

/// <summary>
/// Shared JSON serialization and error handling for all Word MCP tools.
/// </summary>
internal static class WordToolsBase
{
    /// <summary>
    /// JSON options used for every tool response: camelCase, no null values, readable output.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    /// <summary>
    /// Runs a tool action and converts any failure into a structured JSON error response.
    /// </summary>
    /// <param name="tool">Name of the tool, used in the error payload.</param>
    /// <param name="action">Name of the action, used in the error payload.</param>
    /// <param name="operation">The operation to run.</param>
    /// <returns>The serialized result or a serialized error.</returns>
    public static string Execute(string tool, string action, Func<object> operation)
    {
        try
        {
            return Serialize(operation());
        }
#pragma warning disable CA1031 // Tool boundary: every failure must be reported as JSON, not a crash
        catch (Exception ex)
        {
            return Serialize(new
            {
                success = false,
                isError = true,
                tool,
                action,
                errorType = ex is ServiceCommandException { ErrorType: { } reported } ? reported : ex.GetType().Name,
                errorMessage = Describe(ex)
            });
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Serializes a value with the shared tool options.
    /// </summary>
    /// <param name="value">The value to serialize.</param>
    /// <returns>The JSON representation.</returns>
    public static string Serialize(object value)
        => JsonSerializer.Serialize(value, JsonOptions);

    /// <summary>
    /// Resolves the batch for a session id, with a helpful error when the id is missing.
    /// </summary>
    /// <param name="sessionId">The session id supplied by the caller.</param>
    /// <returns>The batch for that session.</returns>
    public static IWordBatch Batch(string? sessionId)
        => ServiceBridge.Batch(sessionId);

    /// <summary>
    /// Validates that a path is an absolute Windows path with a supported extension.
    /// </summary>
    /// <param name="path">The path to validate.</param>
    /// <param name="mustExist">Whether the file has to exist already.</param>
    /// <returns>The normalized full path.</returns>
    public static string ValidatePath(string? path, bool mustExist)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("path is required.", nameof(path));
        }

        if (!Path.IsPathFullyQualified(path))
        {
            throw new ArgumentException(
                $"path must be absolute, for example C:\\Users\\me\\Documents\\report.docx (got '{path}').",
                nameof(path));
        }

        var fullPath = Path.GetFullPath(path);
        var extension = Path.GetExtension(fullPath);

        if (!ComInteropConstants.SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"Unsupported file type '{extension}'. Supported: {string.Join(", ", ComInteropConstants.SupportedExtensions)}.",
                nameof(path));
        }

        if (mustExist && !File.Exists(fullPath))
        {
            throw new FileNotFoundException($"File not found: {fullPath}", fullPath);
        }

        return fullPath;
    }

    /// <summary>
    /// Unwraps a value type an action cannot run without.
    /// </summary>
    /// <typeparam name="T">The underlying value type.</typeparam>
    /// <param name="value">The value supplied by the caller.</param>
    /// <param name="name">Parameter name, used in the error message.</param>
    /// <returns>The value.</returns>
    public static T Require<T>(T? value, string name)
        where T : struct
        => value ?? throw new ArgumentException($"{name} is required for this action.", name);

    /// <summary>
    /// Unwraps a reference type an action cannot run without.
    /// </summary>
    /// <typeparam name="T">The reference type.</typeparam>
    /// <param name="value">The value supplied by the caller.</param>
    /// <param name="name">Parameter name, used in the error message.</param>
    /// <returns>The value.</returns>
    public static T RequireValue<T>(T? value, string name)
        where T : class
        => value ?? throw new ArgumentException($"{name} is required for this action.", name);

    /// <summary>
    /// Unwraps a string an action cannot run without. Blank counts as missing, because a
    /// whitespace-only style or bookmark name is never what the caller meant.
    /// </summary>
    /// <param name="value">The value supplied by the caller.</param>
    /// <param name="name">Parameter name, used in the error message.</param>
    /// <returns>The value.</returns>
    public static string RequireText(string? value, string name)
        => string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException($"{name} is required for this action.", name)
            : value;

    private static string Describe(Exception ex) => ex switch    {
        KeyNotFoundException => ex.Message,
        FileNotFoundException => ex.Message,
        TimeoutException => $"{ex.Message} Word may be showing a dialog; close it and retry.",
        System.Runtime.InteropServices.COMException com =>
            $"Word rejected the operation (HRESULT 0x{com.HResult:X8}): {com.Message}",
        _ => ex.Message
    };
}

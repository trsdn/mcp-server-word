using System.Text.Json;
using WordMcp.Service;

namespace WordMcp.McpServer;

/// <summary>
/// The single seam between the MCP tools and the session service.
/// </summary>
/// <remarks>
/// <para>Before this existed, the file tool and <see cref="WordMcpService"/> each carried their own
/// copy of open, create, save, close, list and test. Two implementations of the same rules drift,
/// and the one the user hits is decided by which entry point they came through. Everything session
/// shaped now goes through here, so there is one answer.</para>
/// <para>Nothing crossing this seam is anything but data. That is what lets the same call run
/// in-process or travel down a named pipe to a daemon, chosen by <c>WORDMCP_SERVICE_MODE</c>,
/// without a single tool knowing which one it got.</para>
/// </remarks>
internal static class ServiceBridge
{
    /// <summary>Environment variable that selects in-process or daemon operation.</summary>
    public const string ModeVariable = "WORDMCP_SERVICE_MODE";

    private static readonly Lock Gate = new();
    private static WordMcpService? _service;
    private static ServiceClient? _client;

    /// <summary>
    /// Gets a value indicating whether commands travel to a separate daemon process.
    /// </summary>
    /// <remarks>
    /// In-process is the default because it is what a single client wants: no second process, no
    /// pipe, no startup wait. Daemon mode is for the case the in-process one cannot serve at all —
    /// several clients that have to see the same open document.
    /// </remarks>
    public static bool UseDaemon
        => string.Equals(
            Environment.GetEnvironmentVariable(ModeVariable),
            "daemon",
            StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Gets the in-process service, creating it on first use so an idle server never starts Word.
    /// </summary>
    public static WordMcpService Service
    {
        get
        {
            lock (Gate)
            {
                return _service ??= new WordMcpService(WordServices.Logger);
            }
        }
    }

    /// <summary>
    /// Runs one service command and returns its payload.
    /// </summary>
    /// <param name="command">The command name, for example <c>session.open</c> or <c>text.insert</c>.</param>
    /// <param name="sessionId">The session the command applies to, when it is session scoped.</param>
    /// <param name="args">Arguments, serialized as the command's argument object.</param>
    /// <returns>The result payload, ready to be serialized into the tool response.</returns>
    /// <exception cref="ServiceCommandException">The command failed.</exception>
    public static object Invoke(string command, string? sessionId = null, object? args = null)
    {
        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args is null ? null : ServiceProtocol.Serialize(args),
            Source = "mcp"
        };

        // The service is asynchronous because a pipe host needs it to be; in-process the only
        // await is the open gate, and blocking on it here is what keeps the tool signatures
        // synchronous for the MCP SDK.
        var response = UseDaemon
            ? Client.SendAsync(request).GetAwaiter().GetResult()
            : Service.ProcessAsync(request).GetAwaiter().GetResult();

        if (!response.Success)
        {
            throw new ServiceCommandException(
                response.ErrorMessage ?? $"Command '{command}' failed.",
                response.ErrorType);
        }

        return response.Result is null
            ? new { success = true }
            : JsonSerializer.Deserialize<JsonElement>(response.Result);
    }

    /// <summary>
    /// Saves and closes every open session. Called on shutdown so unsaved work is not lost.
    /// </summary>
    /// <remarks>
    /// In daemon mode this deliberately leaves the sessions alone: outliving one client is the
    /// entire point of the daemon, and the idle timeout is what ends its life.
    /// </remarks>
    public static void Shutdown()
    {
        lock (Gate)
        {
            _service?.Dispose();
            _service = null;
            _client?.Dispose();
            _client = null;
        }
    }

    private static ServiceClient Client
    {
        get
        {
            lock (Gate)
            {
                return _client ??= new ServiceClient(
                    ServiceSecurity.GetServicePipeName(),
                    logger: WordServices.Logger);
            }
        }
    }
}

/// <summary>
/// A command the service refused or could not complete.
/// </summary>
/// <remarks>
/// The service reports failures as data, not as exceptions, but the tool layer formats errors from
/// exceptions. This carries the original type name across that gap so a caller still sees
/// <c>FileNotFoundException</c> rather than the name of this wrapper.
/// </remarks>
public sealed class ServiceCommandException : Exception
{
    /// <summary>
    /// Creates an exception for a failed command.
    /// </summary>
    /// <param name="message">The failure description reported by the service.</param>
    /// <param name="errorType">The CLR type name the service reported.</param>
    public ServiceCommandException(string message, string? errorType)
        : base(message)
        => ErrorType = errorType;

    /// <summary>
    /// Creates an exception with a message.
    /// </summary>
    /// <param name="message">The failure description.</param>
    public ServiceCommandException(string message)
        : base(message)
    {
    }

    /// <summary>
    /// Creates an exception with a message and an inner exception.
    /// </summary>
    /// <param name="message">The failure description.</param>
    /// <param name="innerException">The underlying failure.</param>
    public ServiceCommandException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    /// <summary>
    /// Creates an exception without a message.
    /// </summary>
    public ServiceCommandException()
    {
    }

    /// <summary>
    /// Gets the CLR type name of the original failure, when the service reported one.
    /// </summary>
    public string? ErrorType { get; }
}

using Microsoft.Extensions.Logging;

namespace WordMcp.McpServer;

/// <summary>
/// Process-wide state used by the MCP tools.
/// </summary>
/// <remarks>
/// The MCP SDK discovers tools as static methods, so what little the tools share cannot be
/// resolved through DI and lives here instead. That is now only the logger: the command
/// implementations moved behind <see cref="ServiceBridge"/>, and so did the sessions, because a
/// tool that holds neither can run against a document in another process.
/// </remarks>
internal static class WordServices
{
    /// <summary>Gets or sets the logger handed to new sessions.</summary>
    public static ILogger? Logger { get; set; }
}

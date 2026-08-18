using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;

namespace WordMcp.ComInterop.Session;

/// <summary>
/// Metadata about an open Word session.
/// </summary>
/// <param name="SessionId">Opaque identifier handed to MCP clients.</param>
/// <param name="FilePath">Full path of the document.</param>
/// <param name="Visible">Whether the Word window is shown.</param>
/// <param name="OpenedAt">UTC timestamp when the session was created.</param>
public sealed record WordSessionInfo(string SessionId, string FilePath, bool Visible, DateTimeOffset OpenedAt);

/// <summary>
/// Tracks open Word sessions so MCP tools can address a document by <c>session_id</c>
/// instead of re-opening it for every call.
/// </summary>
/// <remarks>
/// One instance is shared per process. Sessions are keyed by an opaque id and additionally
/// indexed by normalized file path so callers can find an already open document.
/// </remarks>
public sealed class SessionManager : IDisposable
{
    private readonly ConcurrentDictionary<string, SessionEntry> _sessions = new(StringComparer.Ordinal);
    private readonly ILogger? _logger;
    private int _disposed;

    /// <summary>
    /// Creates a new session manager.
    /// </summary>
    /// <param name="logger">Optional logger passed to every batch created by this manager.</param>
    public SessionManager(ILogger? logger = null) => _logger = logger;

    /// <summary>
    /// Opens an existing document and registers a new session.
    /// </summary>
    /// <param name="filePath">Path to the document.</param>
    /// <param name="visible">Whether the Word window is shown.</param>
    /// <param name="timeoutSeconds">Per-operation timeout in seconds.</param>
    /// <returns>Metadata for the newly created session.</returns>
    public WordSessionInfo Open(string filePath, bool visible = false, int timeoutSeconds = 300)
    {
        var batch = WordSession.BeginBatch(visible, TimeSpan.FromSeconds(timeoutSeconds), _logger, filePath);
        return Register(batch, visible);
    }

    /// <summary>
    /// Creates a new document and registers a session for it.
    /// </summary>
    /// <param name="filePath">Path of the document to create.</param>
    /// <param name="visible">Whether the Word window is shown.</param>
    /// <param name="timeoutSeconds">Per-operation timeout in seconds.</param>
    /// <returns>Metadata for the newly created session.</returns>
    public WordSessionInfo Create(string filePath, bool visible = false, int timeoutSeconds = 300)
    {
        var batch = WordSession.CreateNew(filePath, visible, TimeSpan.FromSeconds(timeoutSeconds), _logger);
        return Register(batch, visible);
    }

    /// <summary>
    /// Resolves the batch for a session id.
    /// </summary>
    /// <param name="sessionId">Session identifier returned by <see cref="Open"/> or <see cref="Create"/>.</param>
    /// <returns>The batch belonging to the session.</returns>
    /// <exception cref="KeyNotFoundException">No session with this id exists.</exception>
    public IWordBatch GetBatch(string sessionId)
    {
        if (!_sessions.TryGetValue(sessionId ?? string.Empty, out var entry))
        {
            throw new KeyNotFoundException(
                $"Unknown session_id '{sessionId}'. Call file(action:'list') to see open sessions, " +
                "or file(action:'open') to start a new one.");
        }

        return entry.Batch;
    }

    /// <summary>
    /// Finds an open session for a document path, if one exists.
    /// </summary>
    /// <param name="filePath">Path to look up.</param>
    /// <returns>The session info, or <c>null</c> when the document is not open.</returns>
    public WordSessionInfo? FindByPath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return null;

        var normalized = Path.GetFullPath(filePath);
        return _sessions.Values
            .FirstOrDefault(e => string.Equals(e.Info.FilePath, normalized, StringComparison.OrdinalIgnoreCase))
            ?.Info;
    }

    /// <summary>
    /// Lists all open sessions.
    /// </summary>
    /// <returns>Metadata for every open session.</returns>
    public IReadOnlyList<WordSessionInfo> List()
        => [.. _sessions.Values.Select(e => e.Info).OrderBy(i => i.OpenedAt)];

    /// <summary>
    /// Saves the document of a session without closing it.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    public void Save(string sessionId) => GetBatch(sessionId).Save();

    /// <summary>
    /// Closes a session, optionally saving the document first.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="save">Whether to save before closing.</param>
    /// <returns><c>true</c> when a session was closed; <c>false</c> when the id was unknown.</returns>
    public bool Close(string sessionId, bool save)
    {
        if (!_sessions.TryRemove(sessionId ?? string.Empty, out var entry))
            return false;

        try
        {
            if (save)
            {
                entry.Batch.Save();
            }
        }
        finally
        {
            entry.Batch.Dispose();
        }

        return true;
    }

    /// <summary>
    /// Closes every open session. Documents are saved before closing to avoid losing work
    /// when the MCP client disconnects.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
            return;

        foreach (var sessionId in _sessions.Keys.ToArray())
        {
            try
            {
                Close(sessionId, save: true);
            }
#pragma warning disable CA1031 // Shutdown must never throw
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "Failed to close Word session {SessionId} during shutdown", sessionId);
            }
#pragma warning restore CA1031
        }
    }

    private WordSessionInfo Register(IWordBatch batch, bool visible)
    {
        var sessionId = $"word-{Guid.NewGuid():N}"[..17];
        var info = new WordSessionInfo(sessionId, batch.DocumentPath, visible, DateTimeOffset.UtcNow);

        if (!_sessions.TryAdd(sessionId, new SessionEntry(batch, info)))
        {
            batch.Dispose();
            throw new InvalidOperationException($"Session id collision for '{sessionId}'.");
        }

        return info;
    }

    private sealed record SessionEntry(IWordBatch Batch, WordSessionInfo Info);
}

using System.Globalization;
using Microsoft.Extensions.Logging;
using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;

namespace WordMcp.Service;

/// <summary>
/// Holds the Word sessions and executes the commands that operate on them.
/// </summary>
/// <remarks>
/// <para>This is the piece that makes a session outlive a single client. It owns the
/// <see cref="SessionManager"/> and exposes it only through <see cref="ProcessAsync"/>, whose
/// request and response types are plain data. Nothing in the contract carries a COM object or a
/// delegate, so the same service can be called in-process or across a named pipe without the
/// callers noticing which one they got.</para>
/// <para>The class is deliberately transport agnostic: it neither opens a pipe nor reads a
/// console. A host does that and calls in here.</para>
/// </remarks>
public sealed class WordMcpService : IDisposable
{
    /// <summary>
    /// The idle period used when the caller does not specify one.
    /// </summary>
    public static readonly TimeSpan DefaultIdleTimeout = TimeSpan.FromMinutes(30);

    private readonly SemaphoreSlim _openGate = new(1, 1);
    private readonly OrphanWordCleanup _cleanup;
    private readonly SessionManager _sessions;
    private readonly TimeProvider _clock;
    private readonly ILogger? _logger;
    private long _lastActivityTicks;
    private int _disposed;

    /// <summary>
    /// Creates a service instance.
    /// </summary>
    /// <param name="logger">Optional logger handed to every session.</param>
    /// <param name="idleTimeout">Idle period after which <see cref="IsIdle"/> reports <c>true</c>.</param>
    /// <param name="clock">Time source; injected by tests so idle behaviour can be checked without waiting.</param>
    public WordMcpService(
        ILogger? logger = null,
        TimeSpan? idleTimeout = null,
        TimeProvider? clock = null)
    {
        _logger = logger;
        _clock = clock ?? TimeProvider.System;
        _cleanup = new OrphanWordCleanup(logger);
        _sessions = new SessionManager(logger);
        IdleTimeout = idleTimeout ?? DefaultIdleTimeout;
        StartedAt = _clock.GetUtcNow();
        _lastActivityTicks = StartedAt.UtcTicks;
    }

    /// <summary>
    /// Gets the UTC time the service was created.
    /// </summary>
    public DateTimeOffset StartedAt { get; }

    /// <summary>
    /// Gets the idle period after which the service may shut itself down.
    /// </summary>
    public TimeSpan IdleTimeout { get; }

    /// <summary>
    /// Gets the UTC time of the most recent request.
    /// </summary>
    public DateTimeOffset LastActivityAt
        => new(Interlocked.Read(ref _lastActivityTicks), TimeSpan.Zero);

    /// <summary>
    /// Gets a value indicating whether a client asked the service to stop.
    /// </summary>
    public bool ShutdownRequested { get; private set; }

    /// <summary>
    /// Gets the number of open sessions.
    /// </summary>
    public int SessionCount => _sessions.List().Count;

    /// <summary>
    /// Gets a value indicating whether the service has been idle long enough to exit.
    /// </summary>
    /// <remarks>
    /// An open session always counts as active, however long ago it was last touched. Closing
    /// Word underneath a client that still holds a session id would turn a dormant workflow into
    /// a broken one, which is worse than a process sitting idle.
    /// </remarks>
    public bool IsIdle
        => SessionCount == 0 && _clock.GetUtcNow() - LastActivityAt >= IdleTimeout;

    /// <summary>
    /// Resolves the batch behind a session id.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <returns>The batch for that session.</returns>
    /// <exception cref="KeyNotFoundException">No session with this id exists.</exception>
    public IWordBatch GetBatch(string? sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ArgumentException(
                "session_id is required. Call file(action:'open', path:'C:\\\\...\\\\document.docx') first.",
                nameof(sessionId));
        }

        return _sessions.GetBatch(sessionId);
    }

    /// <summary>
    /// Runs a single command.
    /// </summary>
    /// <param name="request">The command and its arguments.</param>
    /// <param name="cancellationToken">Token that cancels the command.</param>
    /// <returns>The result, with any failure reported as an unsuccessful response rather than an exception.</returns>
    public async Task<ServiceResponse> ProcessAsync(
        ServiceRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) == 1, this);

        Interlocked.Exchange(ref _lastActivityTicks, _clock.GetUtcNow().UtcTicks);

        try
        {
            return await DispatchAsync(request, cancellationToken).ConfigureAwait(false);
        }
#pragma warning disable CA1031 // Service boundary: a failed command must not take the service down
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Command {Command} failed", request.Command);
            return ServiceResponse.Fail(Describe(ex), ex.GetType().Name);
        }
#pragma warning restore CA1031
        finally
        {
            Interlocked.Exchange(ref _lastActivityTicks, _clock.GetUtcNow().UtcTicks);
        }
    }

    /// <summary>
    /// Saves and closes every session, then terminates any Word process left behind.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
        {
            return;
        }

        // Order matters: the session manager saves and closes documents cleanly, and only what
        // survives that is genuinely orphaned.
        _sessions.Dispose();
        _cleanup.CleanUp();
        _openGate.Dispose();
    }

    private async Task<ServiceResponse> DispatchAsync(ServiceRequest request, CancellationToken cancellationToken)
        => request.Command switch
        {
            "service.ping" => ServiceResponse.Ok(new { pong = true, processId = Environment.ProcessId }),
            "service.status" => ServiceResponse.Ok(BuildStatus()),
            "service.shutdown" => Shutdown(),
            "session.open" => await OpenAsync(request, create: false, cancellationToken).ConfigureAwait(false),
            "session.create" => await OpenAsync(request, create: true, cancellationToken).ConfigureAwait(false),
            "session.save" => Save(request.SessionId),
            "session.close" => Close(request.SessionId, Args(request).Save ?? true),
            "session.list" => List(),
            "session.test" => Test(Args(request).FilePath),
            _ => ServiceResponse.Fail(
                $"Unknown command '{request.Command}'.", nameof(NotSupportedException))
        };

    private static SessionArgs Args(ServiceRequest request)
        => (request.Args is null ? null : ServiceProtocol.Deserialize<SessionArgs>(request.Args)) ?? new SessionArgs();

    private ServiceStatus BuildStatus() => new()
    {
        ProcessId = Environment.ProcessId,
        SessionCount = SessionCount,
        StartedAt = StartedAt,
        LastActivityAt = LastActivityAt,
        IdleTimeout = IdleTimeout,
        Version = typeof(WordMcpService).Assembly.GetName().Version?.ToString() ?? "0.1.0"
    };

    private ServiceResponse Shutdown()
    {
        ShutdownRequested = true;
        return ServiceResponse.Ok(new { success = true, message = "Service is shutting down." });
    }

    private async Task<ServiceResponse> OpenAsync(ServiceRequest request, bool create, CancellationToken cancellationToken)
    {
        var args = Args(request);
        var fullPath = NormalizePath(args.FilePath);
        var visible = args.Visible ?? false;
        var timeout = args.TimeoutSeconds ?? 300;

        // Two clients racing on the same document would otherwise open Word twice and the second
        // would fail on the file lock, so opening is serialized across the whole service.
        await _openGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!create)
            {
                if (!File.Exists(fullPath))
                {
                    throw new FileNotFoundException(
                        $"Word document not found: {fullPath}. Use 'session.create' to create it.", fullPath);
                }

                var existing = _sessions.FindByPath(fullPath);
                if (existing != null)
                {
                    return ServiceResponse.Ok(new
                    {
                        success = true,
                        sessionId = existing.SessionId,
                        filePath = existing.FilePath,
                        visible = existing.Visible,
                        reused = true,
                        message = "Document is already open; reusing the existing session."
                    });
                }
            }
            else if (File.Exists(fullPath))
            {
                throw new IOException(
                    $"File already exists: {fullPath}. Use 'session.open' instead, or choose a different path.");
            }

            var info = create
                ? _sessions.Create(fullPath, visible, timeout)
                : _sessions.Open(fullPath, visible, timeout);

            _cleanup.Track(_sessions.GetBatch(info.SessionId).WordProcessId);

            return ServiceResponse.Ok(new
            {
                success = true,
                sessionId = info.SessionId,
                filePath = info.FilePath,
                visible = info.Visible,
                reused = false,
                message = create
                    ? "Document created and session opened."
                    : "Session opened. Pass session_id to the other tools."
            });
        }
        finally
        {
            _openGate.Release();
        }
    }

    private ServiceResponse Save(string? sessionId)
    {
        GetBatch(sessionId).Save();
        return ServiceResponse.Ok(new { success = true, sessionId, message = "Document saved." });
    }

    private ServiceResponse Close(string? sessionId, bool save)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ArgumentException("session_id is required for 'session.close'.", nameof(sessionId));
        }

        // Read the process id before closing; afterwards the batch is gone.
        int? wordProcessId = null;
        try
        {
            wordProcessId = _sessions.GetBatch(sessionId).WordProcessId;
        }
        catch (KeyNotFoundException)
        {
            // Reported through the result below.
        }

        var closed = _sessions.Close(sessionId, save);
        _cleanup.Forget(wordProcessId);

        return ServiceResponse.Ok(new
        {
            success = closed,
            sessionId,
            saved = save && closed,
            message = closed
                ? save ? "Document saved and session closed." : "Session closed without saving."
                : $"No open session with id '{sessionId}'."
        });
    }

    private ServiceResponse List()
    {
        var sessions = _sessions.List()
            .Select(s => new
            {
                sessionId = s.SessionId,
                filePath = s.FilePath,
                visible = s.Visible,
                openedAt = s.OpenedAt
            })
            .ToArray();

        return ServiceResponse.Ok(new { success = true, count = sessions.Length, sessions });
    }

    private static ServiceResponse Test(string? filePath)
    {
        var fullPath = NormalizePath(filePath);
        var extension = Path.GetExtension(fullPath);
        var exists = File.Exists(fullPath);
        var problems = new List<string>();

        if (!exists)
        {
            problems.Add("File does not exist.");
        }

        if (!ComInteropConstants.SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            problems.Add(string.Create(CultureInfo.InvariantCulture, $"Unsupported extension '{extension}'."));
        }

        if (exists)
        {
            try
            {
                FileAccessValidator.ValidateFileNotLocked(fullPath);
            }
            catch (InvalidOperationException ex)
            {
                problems.Add(ex.Message);
            }

            if (FileAccessValidator.IsIrmProtected(fullPath))
            {
                problems.Add("File appears to be IRM/RMS protected and cannot be automated.");
            }
        }

        return ServiceResponse.Ok(new
        {
            success = problems.Count == 0,
            filePath = fullPath,
            exists,
            extension,
            canOpen = problems.Count == 0,
            problems,
            message = problems.Count == 0 ? "File can be opened." : string.Join(" ", problems)
        });
    }

    private static string NormalizePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("filePath is required for this command.", nameof(path));
        }

        if (!Path.IsPathFullyQualified(path))
        {
            throw new ArgumentException(
                $"filePath must be absolute, for example C:\\Users\\me\\Documents\\report.docx (got '{path}').",
                nameof(path));
        }

        return Path.GetFullPath(path);
    }

    private static string Describe(Exception ex) => ex switch
    {
        TimeoutException => $"{ex.Message} Word may be showing a dialog; close it and retry.",
        System.Runtime.InteropServices.COMException com =>
            $"Word rejected the operation (HRESULT 0x{com.HResult:X8}): {com.Message}",
        _ => ex.Message
    };

    /// <summary>
    /// Arguments accepted by the <c>session.*</c> commands.
    /// </summary>
    private sealed class SessionArgs
    {
        public string? FilePath { get; init; }

        public bool? Visible { get; init; }

        public bool? Save { get; init; }

        public int? TimeoutSeconds { get; init; }
    }
}

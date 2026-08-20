using System.Text;
using Microsoft.Extensions.Logging;

namespace WordMcp.Service;

/// <summary>
/// Serves a <see cref="WordMcpService"/> over a named pipe.
/// </summary>
/// <remarks>
/// <para>Each accepted connection is handled on its own task and may carry many requests, so a
/// client can hold the pipe open for a whole workflow. A fresh listening instance is created
/// before the previous connection is handed off, which keeps a second client from being refused
/// while the first is busy.</para>
/// <para>The host stops when a client sends <c>service.shutdown</c>, when the service has been
/// idle past its timeout, or when the caller cancels. Requests are not serialized here: the
/// service itself decides what needs a lock, and Word calls are already serialized per session on
/// that session's STA thread.</para>
/// </remarks>
public sealed class ServiceHost : IDisposable
{
    private readonly WordMcpService _service;
    private readonly ILogger? _logger;
    private readonly Mutex? _singleInstance;
    private int _disposed;

    private ServiceHost(WordMcpService service, string pipeName, Mutex? singleInstance, ILogger? logger)
    {
        _service = service;
        _logger = logger;
        _singleInstance = singleInstance;
        PipeName = pipeName;
    }

    /// <summary>
    /// Gets the pipe this host listens on.
    /// </summary>
    public string PipeName { get; }

    /// <summary>
    /// How often the host checks whether it should stop.
    /// </summary>
    public static TimeSpan StopPollInterval { get; } = TimeSpan.FromSeconds(1);

    /// <summary>
    /// Claims the pipe name for this process and creates a host for it.
    /// </summary>
    /// <param name="service">The service whose commands are served.</param>
    /// <param name="pipeName">Name of the pipe to listen on.</param>
    /// <param name="logger">Optional logger for diagnostics.</param>
    /// <returns>The host, or <c>null</c> when another process already serves this pipe.</returns>
    /// <remarks>
    /// A named mutex, not the pipe itself, decides who wins: Windows happily allows several server
    /// instances on one pipe name, so two daemons would otherwise both appear to start and clients
    /// would be split between them, each seeing half the sessions.
    /// </remarks>
    public static ServiceHost? TryCreate(WordMcpService service, string pipeName, ILogger? logger = null)
    {
        ArgumentNullException.ThrowIfNull(service);
        ArgumentException.ThrowIfNullOrWhiteSpace(pipeName);

        var mutex = new Mutex(initiallyOwned: true, $"WordMcp-host-{pipeName}", out var createdNew);
        if (!createdNew)
        {
            mutex.Dispose();
            return null;
        }

        return new ServiceHost(service, pipeName, mutex, logger);
    }

    /// <summary>
    /// Accepts connections until the service asks to stop or the caller cancels.
    /// </summary>
    /// <param name="cancellationToken">Token that stops the host.</param>
    /// <returns>A task that completes once the host has stopped listening.</returns>
    public async Task RunAsync(CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) == 1, this);

        using var stopping = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var watching = WatchForStopAsync(stopping);
        var connections = new List<Task>();

        _logger?.LogInformation("Word session service listening on {PipeName}", PipeName);

        try
        {
            while (!stopping.IsCancellationRequested)
            {
                var pipe = ServiceSecurity.CreateSecureServer(PipeName);

                try
                {
                    await pipe.WaitForConnectionAsync(stopping.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    await pipe.DisposeAsync().ConfigureAwait(false);
                    break;
                }
#pragma warning disable CA1031 // A failed accept must not end the host
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Could not accept a connection on {PipeName}", PipeName);
                    await pipe.DisposeAsync().ConfigureAwait(false);
                    continue;
                }
#pragma warning restore CA1031

                connections.Add(ServeAsync(pipe, stopping.Token));
                connections.RemoveAll(t => t.IsCompleted);
            }
        }
        finally
        {
            await stopping.CancelAsync().ConfigureAwait(false);
            await watching.ConfigureAwait(false);

            // Let in-flight requests finish so a client never sees a half-written response.
            await Task.WhenAll(connections).ConfigureAwait(false);
            _logger?.LogInformation("Word session service on {PipeName} stopped", PipeName);
        }
    }

    /// <summary>
    /// Releases the claim on the pipe name.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
        {
            return;
        }

        if (_singleInstance is not null)
        {
            try
            {
                _singleInstance.ReleaseMutex();
            }
            catch (ApplicationException)
            {
                // Not owned any more; nothing to release.
            }

            _singleInstance.Dispose();
        }
    }

    private async Task WatchForStopAsync(CancellationTokenSource stopping)
    {
        try
        {
            while (!stopping.IsCancellationRequested)
            {
                await Task.Delay(StopPollInterval, stopping.Token).ConfigureAwait(false);

                if (_service.ShutdownRequested || _service.IsIdle)
                {
                    _logger?.LogInformation(
                        "Stopping: {Reason}",
                        _service.ShutdownRequested ? "a client requested shutdown" : "idle timeout elapsed");

                    // Nudge the accept loop, which is otherwise parked on WaitForConnectionAsync.
                    await using var nudge = ServiceSecurity.CreateClient(PipeName);
                    await stopping.CancelAsync().ConfigureAwait(false);

                    try
                    {
                        await nudge.ConnectAsync(500, CancellationToken.None).ConfigureAwait(false);
                    }
                    catch (TimeoutException)
                    {
                        // The loop was not waiting; it will notice the cancellation itself.
                    }

                    return;
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown.
        }
    }

    private async Task ServeAsync(System.IO.Pipes.NamedPipeServerStream pipe, CancellationToken cancellationToken)
    {
        await Task.Yield();

        try
        {
            // No BOM and no auto-flush games: the framing is one line per message, so the writer
            // flushes explicitly after each response.
            using var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, leaveOpen: true);
            await using var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, leaveOpen: true);

            while (!cancellationToken.IsCancellationRequested && pipe.IsConnected)
            {
                var line = await reader.ReadLineAsync(cancellationToken).ConfigureAwait(false);
                if (line is null)
                {
                    break;
                }

                var response = await HandleAsync(line, cancellationToken).ConfigureAwait(false);

                await writer.WriteLineAsync(ServiceProtocol.Serialize(response).AsMemory(), cancellationToken)
                    .ConfigureAwait(false);
                await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // The host is stopping.
        }
        catch (IOException)
        {
            // The client went away mid-conversation; nothing to report.
        }
#pragma warning disable CA1031 // One bad connection must not take the host down
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Connection on {PipeName} failed", PipeName);
        }
#pragma warning restore CA1031
        finally
        {
            await pipe.DisposeAsync().ConfigureAwait(false);
        }
    }

    private async Task<ServiceResponse> HandleAsync(string line, CancellationToken cancellationToken)
    {
        ServiceRequest? request;

        try
        {
            request = ServiceProtocol.Deserialize<ServiceRequest>(line);
        }
        catch (System.Text.Json.JsonException ex)
        {
            return ServiceResponse.Fail($"Malformed request: {ex.Message}", nameof(System.Text.Json.JsonException));
        }

        return request is null
            ? ServiceResponse.Fail("Empty request.", nameof(ArgumentException))
            : await _service.ProcessAsync(request, cancellationToken).ConfigureAwait(false);
    }
}

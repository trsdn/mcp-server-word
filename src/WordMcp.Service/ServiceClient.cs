using System.Diagnostics;
using System.Globalization;
using System.Text;
using Microsoft.Extensions.Logging;

namespace WordMcp.Service;

/// <summary>
/// Options that control how a <see cref="ServiceClient"/> reaches its daemon.
/// </summary>
public sealed class ServiceClientOptions
{
    /// <summary>
    /// Gets the time allowed for a single connection attempt.
    /// </summary>
    public TimeSpan ConnectTimeout { get; init; } = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Gets the time allowed for a whole request, including waiting for Word.
    /// </summary>
    /// <remarks>
    /// Generous on purpose. A Word operation carries its own timeout, and the pipe expiring first
    /// would replace a precise error with a vague one.
    /// </remarks>
    public TimeSpan RequestTimeout { get; init; } = TimeSpan.FromHours(2);

    /// <summary>
    /// Gets a value indicating whether the client may start a daemon that is not running.
    /// </summary>
    public bool AutoStart { get; init; } = true;

    /// <summary>
    /// Gets the time to wait for a freshly started daemon to answer.
    /// </summary>
    public TimeSpan StartupTimeout { get; init; } = TimeSpan.FromSeconds(20);

    /// <summary>
    /// Gets the daemon executable. Probed next to the running assembly when not set.
    /// </summary>
    public string? ExecutablePath { get; init; }

    /// <summary>
    /// Gets the idle timeout handed to a daemon this client starts.
    /// </summary>
    public TimeSpan? IdleTimeout { get; init; }
}

/// <summary>
/// Talks to a <see cref="ServiceHost"/> over a named pipe, starting the daemon when needed.
/// </summary>
/// <remarks>
/// <para>Every request gets its own connection. That costs a few hundred microseconds on local
/// IPC and buys reconnection for free: a daemon that exited on its idle timeout is simply started
/// again by the next call, with no stale-socket state to reason about.</para>
/// <para>Transport failures are returned as unsuccessful responses rather than thrown, so callers
/// handle "the service is gone" the same way they handle "the command failed".</para>
/// </remarks>
public sealed class ServiceClient : IDisposable
{
    private readonly SemaphoreSlim _startGate = new(1, 1);
    private readonly ServiceClientOptions _options;
    private readonly ILogger? _logger;
    private int _disposed;

    /// <summary>
    /// Creates a client for a pipe.
    /// </summary>
    /// <param name="pipeName">Name of the pipe, usually <see cref="ServiceSecurity.GetServicePipeName"/>.</param>
    /// <param name="options">Connection and auto-start behaviour.</param>
    /// <param name="logger">Optional logger for diagnostics.</param>
    public ServiceClient(string pipeName, ServiceClientOptions? options = null, ILogger? logger = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pipeName);

        PipeName = pipeName;
        _options = options ?? new ServiceClientOptions();
        _logger = logger;
    }

    /// <summary>
    /// Gets the pipe this client talks to.
    /// </summary>
    public string PipeName { get; }

    /// <summary>
    /// Sends one command and waits for its result.
    /// </summary>
    /// <param name="request">The command to run.</param>
    /// <param name="cancellationToken">Token that cancels the call.</param>
    /// <returns>The service's response, or a failed response describing the transport problem.</returns>
    public async Task<ServiceResponse> SendAsync(ServiceRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) == 1, this);

        var response = await TrySendAsync(request, cancellationToken).ConfigureAwait(false);
        if (response is not null)
        {
            return response;
        }

        if (!_options.AutoStart)
        {
            return ServiceResponse.Fail(
                $"No Word session service is listening on '{PipeName}'.", nameof(IOException));
        }

        if (!await EnsureRunningAsync(cancellationToken).ConfigureAwait(false))
        {
            return ServiceResponse.Fail(
                "Could not start the Word session service.", nameof(InvalidOperationException));
        }

        return await TrySendAsync(request, cancellationToken).ConfigureAwait(false)
            ?? ServiceResponse.Fail(
                "The Word session service started but did not answer.", nameof(TimeoutException));
    }

    /// <summary>
    /// Checks whether a service is answering on the pipe. Never starts one.
    /// </summary>
    /// <param name="cancellationToken">Token that cancels the call.</param>
    /// <returns><c>true</c> when a service answered.</returns>
    public async Task<bool> PingAsync(CancellationToken cancellationToken = default)
    {
        var response = await TrySendAsync(
            new ServiceRequest { Command = "service.ping", Source = "client" },
            cancellationToken).ConfigureAwait(false);

        return response?.Success == true;
    }

    /// <summary>
    /// Makes sure a daemon is running, starting one if necessary.
    /// </summary>
    /// <param name="cancellationToken">Token that cancels the call.</param>
    /// <returns><c>true</c> when a service is answering by the time this returns.</returns>
    public async Task<bool> EnsureRunningAsync(CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) == 1, this);

        if (await PingAsync(cancellationToken).ConfigureAwait(false))
        {
            return true;
        }

        // Two callers racing here would spawn two daemons; the loser's would exit on the single
        // instance mutex, but starting Word twice is expensive enough to be worth avoiding.
        await _startGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (await PingAsync(cancellationToken).ConfigureAwait(false))
            {
                return true;
            }

            if (!TryStartDaemon())
            {
                return false;
            }

            return await WaitForServiceAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _startGate.Release();
        }
    }

    /// <summary>
    /// Locates the daemon executable.
    /// </summary>
    /// <returns>The full path, or <c>null</c> when it could not be found.</returns>
    public static string? FindExecutable()
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "WordMcp.Service.exe"),
            Path.Combine(Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory, "WordMcp.Service.exe")
        };

        return Array.Find(candidates, File.Exists);
    }

    /// <summary>
    /// Releases the client's resources. The daemon keeps running.
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) == 0)
        {
            _startGate.Dispose();
        }
    }

    /// <summary>
    /// Performs one round trip.
    /// </summary>
    /// <returns>The response, or <c>null</c> when no service could be reached.</returns>
    private async Task<ServiceResponse?> TrySendAsync(ServiceRequest request, CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(_options.RequestTimeout);

        try
        {
            await using var pipe = ServiceSecurity.CreateClient(PipeName);
            await pipe.ConnectAsync((int)_options.ConnectTimeout.TotalMilliseconds, timeout.Token)
                .ConfigureAwait(false);

            await using var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, leaveOpen: true);
            await writer.WriteLineAsync(ServiceProtocol.Serialize(request).AsMemory(), timeout.Token)
                .ConfigureAwait(false);
            await writer.FlushAsync(timeout.Token).ConfigureAwait(false);

            using var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, leaveOpen: true);
            var line = await reader.ReadLineAsync(timeout.Token).ConfigureAwait(false);

            if (line is null)
            {
                return ServiceResponse.Fail(
                    "The Word session service closed the connection without answering.", nameof(IOException));
            }

            return ServiceProtocol.Deserialize<ServiceResponse>(line)
                ?? ServiceResponse.Fail("The Word session service sent an empty answer.", nameof(IOException));
        }
        catch (TimeoutException)
        {
            // No listener on the pipe. Reported as "unreachable" so the caller can start one.
            return null;
        }
        catch (OperationCanceledException) when (timeout.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
        {
            return ServiceResponse.Fail(
                string.Create(CultureInfo.InvariantCulture, $"The request timed out after {_options.RequestTimeout}."),
                nameof(TimeoutException));
        }
        catch (IOException)
        {
            // The daemon died between connect and answer; treat it as unreachable and retry.
            return null;
        }
        catch (UnauthorizedAccessException ex)
        {
            return ServiceResponse.Fail(
                $"Access to the Word session service was denied: {ex.Message}", nameof(UnauthorizedAccessException));
        }
    }

    private bool TryStartDaemon()
    {
        var executable = _options.ExecutablePath ?? FindExecutable();
        if (executable is null)
        {
            _logger?.LogWarning("WordMcp.Service.exe was not found next to the running application");
            return false;
        }

        var startInfo = new ProcessStartInfo(executable)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        startInfo.ArgumentList.Add("--daemon");
        startInfo.ArgumentList.Add("--pipe");
        startInfo.ArgumentList.Add(PipeName);

        if (_options.IdleTimeout is { } idle)
        {
            startInfo.ArgumentList.Add("--idle-minutes");
            startInfo.ArgumentList.Add(((int)idle.TotalMinutes).ToString(CultureInfo.InvariantCulture));
        }

        try
        {
            using var process = Process.Start(startInfo);
            _logger?.LogInformation("Started the Word session service (pid {ProcessId})", process?.Id);
            return process is not null;
        }
#pragma warning disable CA1031 // A failed start is reported to the caller, not thrown
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Could not start {Executable}", executable);
            return false;
        }
#pragma warning restore CA1031
    }

    private async Task<bool> WaitForServiceAsync(CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + _options.StartupTimeout;

        while (DateTimeOffset.UtcNow < deadline)
        {
            if (await PingAsync(cancellationToken).ConfigureAwait(false))
            {
                return true;
            }

            await Task.Delay(100, cancellationToken).ConfigureAwait(false);
        }

        return false;
    }
}

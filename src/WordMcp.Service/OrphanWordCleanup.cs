using System.Collections.Concurrent;
using System.Diagnostics;
using Microsoft.Extensions.Logging;

namespace WordMcp.Service;

/// <summary>
/// Terminates <c>WINWORD.EXE</c> processes that this service started and that outlived their session.
/// </summary>
/// <remarks>
/// <para>Word normally exits when the last automation client releases it, but a document that is
/// left in a modal state — a repair prompt, a recovery pane — keeps the process alive with no
/// window anyone can reach. Over a long-running service those accumulate and eventually lock the
/// very documents the next session wants to open.</para>
/// <para>Only process ids handed to <see cref="Track"/> are ever considered. A Word instance the
/// user started by hand is never seen by this class and therefore never killed, which is the whole
/// point of tracking rather than scanning by process name.</para>
/// </remarks>
public sealed class OrphanWordCleanup
{
    private readonly ConcurrentDictionary<int, DateTimeOffset> _tracked = new();
    private readonly ILogger? _logger;

    /// <summary>
    /// Creates a cleanup registry.
    /// </summary>
    /// <param name="logger">Optional logger for diagnostics.</param>
    public OrphanWordCleanup(ILogger? logger = null) => _logger = logger;

    /// <summary>
    /// Gets the process ids currently being tracked.
    /// </summary>
    public IReadOnlyCollection<int> TrackedProcessIds => [.. _tracked.Keys];

    /// <summary>
    /// Records a Word process id belonging to a live session.
    /// </summary>
    /// <param name="processId">The process id, or <c>null</c> when it could not be determined.</param>
    public void Track(int? processId)
    {
        if (processId is > 0)
        {
            _tracked[processId.Value] = DateTimeOffset.UtcNow;
        }
    }

    /// <summary>
    /// Stops tracking a process id without terminating it.
    /// </summary>
    /// <param name="processId">The process id to forget.</param>
    public void Forget(int? processId)
    {
        if (processId is > 0)
        {
            _tracked.TryRemove(processId.Value, out _);
        }
    }

    /// <summary>
    /// Terminates every tracked process that is still running and forgets the rest.
    /// </summary>
    /// <returns>The number of processes that had to be terminated.</returns>
    public int CleanUp()
    {
        var killed = 0;

        foreach (var processId in _tracked.Keys.ToArray())
        {
            _tracked.TryRemove(processId, out _);

            if (TryKill(processId))
            {
                killed++;
            }
        }

        return killed;
    }

    private bool TryKill(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);

            // A pid is reused once the original process is gone, so terminating by number alone
            // could hit an unrelated program. Only a live WINWORD is a candidate.
            if (process.HasExited ||
                !string.Equals(process.ProcessName, "WINWORD", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            process.Kill(entireProcessTree: false);
            _logger?.LogWarning("Terminated orphaned Word process {ProcessId}", processId);
            return true;
        }
        catch (ArgumentException)
        {
            // Already gone — the normal case.
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
#pragma warning disable CA1031 // Cleanup runs during shutdown and must never throw
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Could not terminate Word process {ProcessId}", processId);
            return false;
        }
#pragma warning restore CA1031
    }
}

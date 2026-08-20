using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace WordMcp.ComInterop;

/// <summary>
/// OLE message filter that handles Word COM busy/retry scenarios.
/// Automatically retries when Word returns RPC_E_SERVERCALL_RETRYLATER.
/// </summary>
/// <remarks>
/// Register once per STA thread via <see cref="Register"/> and revoke on thread shutdown
/// via <see cref="Revoke"/>. Without this filter, transient "server busy" states surface as
/// COMException instead of being retried.
/// </remarks>
[GeneratedComClass]
public sealed partial class OleMessageFilter : IOleMessageFilter
{
    private static readonly StrategyBasedComWrappers s_comWrappers = new();

    /// <summary>Maximum time a single COM call is retried while Word reports itself busy.</summary>
    private const int RetryTimeoutMs = 30_000;

    /// <summary>
    /// Idle time after which the next rejection is treated as a new call rather than a
    /// continuation of the previous retry sequence.
    /// </summary>
    private const int RetrySequenceResetMs = 2_000;

    [ThreadStatic]
    private static long _retrySequenceStart;

    [ThreadStatic]
    private static long _lastRetryTick;

    [ThreadStatic]
    private static nint _oldFilterPtr;

    [ThreadStatic]
    private static bool _isRegistered;

    /// <summary>
    /// Registers the OLE message filter for the current STA thread.
    /// </summary>
    /// <exception cref="InvalidOperationException">Already registered on this thread, or registration failed.</exception>
    public static void Register()
    {
        if (_isRegistered)
        {
            throw new InvalidOperationException("OLE message filter is already registered on this thread.");
        }

        var newFilter = new OleMessageFilter();
        nint newFilterPtr = s_comWrappers.GetOrCreateComInterfaceForObject(newFilter, CreateComInterfaceFlags.None);

        int result = CoRegisterMessageFilter(newFilterPtr, out _oldFilterPtr);
        if (result != 0)
        {
            throw new InvalidOperationException($"Failed to register OLE message filter. HRESULT: 0x{result:X8}");
        }

        _isRegistered = true;
    }

    /// <summary>
    /// Revokes the OLE message filter and restores the previous filter.
    /// Safe to call when <see cref="Register"/> was never called.
    /// </summary>
    public static void Revoke()
    {
        if (!_isRegistered)
        {
            return;
        }

        _ = CoRegisterMessageFilter(_oldFilterPtr, out _);

        _oldFilterPtr = 0;
        _isRegistered = false;
    }

    /// <summary>
    /// Gets a value indicating whether the filter is registered on the current thread.
    /// </summary>
    public static bool IsRegistered => _isRegistered;

    /// <summary>
    /// Handles incoming COM calls. Always accepts them (SERVERCALL_ISHANDLED).
    /// </summary>
    int IOleMessageFilter.HandleInComingCall(int dwCallType, nint htaskCaller, int dwTickCount, nint lpInterfaceInfo)
        => 0; // SERVERCALL_ISHANDLED

    /// <summary>
    /// Handles rejected COM calls with exponential backoff while Word is busy.
    /// </summary>
    /// <returns>Milliseconds to wait before retrying, or -1 to cancel the call.</returns>
    /// <remarks>
    /// Word rejects calls with <c>SERVERCALL_REJECTED</c> for several seconds after startup,
    /// so both rejection kinds have to be retried. <c>dwTickCount</c> is only meaningful for
    /// <c>SERVERCALL_RETRYLATER</c>, so the retry deadline is tracked per thread instead.
    /// </remarks>
    int IOleMessageFilter.RetryRejectedCall(nint htaskCallee, int dwTickCount, int dwRejectType)
    {
        const int ServerCallRejected = 1;
        const int ServerCallRetryLater = 2;

        if (dwRejectType is not (ServerCallRejected or ServerCallRetryLater))
        {
            return -1; // Cancel immediately for non-retryable rejections
        }

        long now = Environment.TickCount64;

        // A gap longer than the reset window means this is a new call, not a continued retry.
        if (_retrySequenceStart == 0 || now - _lastRetryTick > RetrySequenceResetMs)
        {
            _retrySequenceStart = now;
        }

        _lastRetryTick = now;
        long elapsed = now - _retrySequenceStart;

        if (elapsed >= RetryTimeoutMs)
        {
            _retrySequenceStart = 0;
            return -1; // Give up
        }

        return elapsed switch
        {
            < 1000 => 100,
            < 5000 => 200,
            < 15000 => 500,
            _ => 1000
        };
    }

    /// <summary>
    /// Handles pending messages during an outgoing COM call by dispatching them
    /// (PENDINGMSG_WAITDEFPROCESS), which keeps embedded OLE activation working.
    /// </summary>
    int IOleMessageFilter.MessagePending(nint htaskCallee, int dwTickCount, int dwPendingType)
        => 2; // PENDINGMSG_WAITDEFPROCESS

    [LibraryImport("Ole32.dll")]
    private static partial int CoRegisterMessageFilter(
        nint lpMessageFilter,
        out nint lplpMessageFilter);
}

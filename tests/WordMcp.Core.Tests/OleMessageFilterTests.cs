using WordMcp.ComInterop;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Covers the retry policy of the OLE message filter, which is plain .NET logic and therefore
/// testable without Word. The COM registration itself is not exercised here.
/// </summary>
public class OleMessageFilterTests
{
    private static readonly IOleMessageFilter Filter = new OleMessageFilter();

    private const int ServerCallRejected = 1;
    private const int ServerCallRetryLater = 2;
    private const int Cancel = -1;

    [Theory]
    [InlineData(ServerCallRejected)]
    [InlineData(ServerCallRetryLater)]
    public void RetryRejectedCall_RetriesBothRejectionKinds(int rejectType)
    {
        // Word rejects calls with SERVERCALL_REJECTED for seconds after startup; treating that as
        // fatal made every session fail to start.
        int delay = Filter.RetryRejectedCall(htaskCallee: 0, dwTickCount: 0, rejectType);

        Assert.True(delay > 0, $"Expected a retry delay but the call was cancelled ({delay}).");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(3)]
    [InlineData(99)]
    public void RetryRejectedCall_CancelsUnknownRejectionKinds(int rejectType)
    {
        Assert.Equal(Cancel, Filter.RetryRejectedCall(htaskCallee: 0, dwTickCount: 0, rejectType));
    }

    [Fact]
    public void RetryRejectedCall_BacksOffAsTheSequenceContinues()
    {
        int first = Filter.RetryRejectedCall(0, 0, ServerCallRetryLater);
        Thread.Sleep(1100);
        int later = Filter.RetryRejectedCall(0, 0, ServerCallRetryLater);

        Assert.True(later > first, $"Expected backoff to grow but got {first} then {later}.");
    }

    [Fact]
    public void HandleInComingCall_AcceptsCalls()
    {
        Assert.Equal(0, Filter.HandleInComingCall(0, 0, 0, 0));
    }

    [Fact]
    public void MessagePending_DispatchesMessages()
    {
        Assert.Equal(2, Filter.MessagePending(0, 0, 0));
    }
}

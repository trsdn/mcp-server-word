using System.Diagnostics;
using Xunit;

namespace WordMcp.Service.Tests;

/// <summary>
/// Orphan cleanup. The dangerous failure mode here is killing something that is not ours, so the
/// central test tracks this very test process and asserts that cleanup leaves it alone.
/// </summary>
public class OrphanWordCleanupTests
{
    [Fact]
    public void Track_RecordsAProcessId()
    {
        var cleanup = new OrphanWordCleanup();

        cleanup.Track(1234);

        Assert.Contains(1234, cleanup.TrackedProcessIds);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0)]
    [InlineData(-1)]
    public void Track_IgnoresAMissingOrNonsensicalProcessId(int? processId)
    {
        var cleanup = new OrphanWordCleanup();

        cleanup.Track(processId);

        Assert.Empty(cleanup.TrackedProcessIds);
    }

    [Fact]
    public void Track_IsIdempotent()
    {
        var cleanup = new OrphanWordCleanup();

        cleanup.Track(1234);
        cleanup.Track(1234);

        Assert.Single(cleanup.TrackedProcessIds);
    }

    [Fact]
    public void Forget_RemovesTheProcessIdWithoutTerminatingIt()
    {
        var cleanup = new OrphanWordCleanup();
        cleanup.Track(Environment.ProcessId);

        cleanup.Forget(Environment.ProcessId);

        Assert.Empty(cleanup.TrackedProcessIds);
        Assert.False(Process.GetCurrentProcess().HasExited);
    }

    [Fact]
    public void Forget_IgnoresAnUnknownProcessId()
    {
        var cleanup = new OrphanWordCleanup();
        cleanup.Track(1234);

        cleanup.Forget(9999);

        Assert.Contains(1234, cleanup.TrackedProcessIds);
    }

    [Fact]
    public void CleanUp_SparesAProcessThatIsNotWord()
    {
        var cleanup = new OrphanWordCleanup();
        cleanup.Track(Environment.ProcessId);

        var killed = cleanup.CleanUp();

        Assert.Equal(0, killed);
        Assert.False(Process.GetCurrentProcess().HasExited);
    }

    [Fact]
    public void CleanUp_IgnoresAProcessThatIsAlreadyGone()
    {
        var cleanup = new OrphanWordCleanup();
        cleanup.Track(int.MaxValue - 1);

        Assert.Equal(0, cleanup.CleanUp());
    }

    [Fact]
    public void CleanUp_ClearsTheRegistry()
    {
        var cleanup = new OrphanWordCleanup();
        cleanup.Track(Environment.ProcessId);
        cleanup.Track(int.MaxValue - 1);

        cleanup.CleanUp();

        Assert.Empty(cleanup.TrackedProcessIds);
    }

    [Fact]
    public void CleanUp_CanBeCalledTwice()
    {
        var cleanup = new OrphanWordCleanup();
        cleanup.Track(Environment.ProcessId);

        cleanup.CleanUp();

        Assert.Equal(0, cleanup.CleanUp());
    }
}

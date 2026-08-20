using System.Text.Json;
using Xunit;

namespace WordMcp.Service.Tests;

/// <summary>
/// Service behaviour that does not need Word: dispatch, argument validation, status and the idle
/// rule. Word-backed lifecycle checks live in <c>SessionLifecycleIntegrationTests</c>.
/// </summary>
public class WordMcpServiceTests
{
    private static async Task<ServiceResponse> RunAsync(WordMcpService service, string command, object? args = null, string? sessionId = null)
        => await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args is null ? null : ServiceProtocol.Serialize(args)
        });

    private static JsonElement Payload(ServiceResponse response)
    {
        Assert.True(response.Success, response.ErrorMessage);
        Assert.NotNull(response.Result);
        return JsonDocument.Parse(response.Result).RootElement.Clone();
    }

    [Fact]
    public async Task Ping_ReportsTheOwningProcess()
    {
        using var service = new WordMcpService();

        var payload = Payload(await RunAsync(service, "service.ping"));

        Assert.True(payload.GetProperty("pong").GetBoolean());
        Assert.Equal(Environment.ProcessId, payload.GetProperty("processId").GetInt32());
    }

    [Fact]
    public async Task Status_ReportsNoSessionsOnAFreshService()
    {
        using var service = new WordMcpService(idleTimeout: TimeSpan.FromMinutes(7));

        var status = ServiceProtocol.Deserialize<ServiceStatus>(
            (await RunAsync(service, "service.status")).Result!);

        Assert.NotNull(status);
        Assert.Equal(0, status.SessionCount);
        Assert.Equal(Environment.ProcessId, status.ProcessId);
        Assert.Equal(TimeSpan.FromMinutes(7), status.IdleTimeout);
        Assert.False(string.IsNullOrWhiteSpace(status.Version));
    }

    [Fact]
    public async Task Shutdown_SetsTheShutdownFlag()
    {
        using var service = new WordMcpService();
        Assert.False(service.ShutdownRequested);

        Assert.True((await RunAsync(service, "service.shutdown")).Success);

        Assert.True(service.ShutdownRequested);
    }

    [Fact]
    public async Task AnUnknownCommandFailsWithoutThrowing()
    {
        using var service = new WordMcpService();

        var response = await RunAsync(service, "session.teleport");

        Assert.False(response.Success);
        Assert.Contains("session.teleport", response.ErrorMessage!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task List_IsEmptyOnAFreshService()
    {
        using var service = new WordMcpService();

        var payload = Payload(await RunAsync(service, "session.list"));

        Assert.Equal(0, payload.GetProperty("count").GetInt32());
        Assert.Empty(payload.GetProperty("sessions").EnumerateArray());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("report.docx")]
    [InlineData("..\\report.docx")]
    public async Task Open_RejectsAPathThatIsNotAbsolute(string? path)
    {
        using var service = new WordMcpService();

        var response = await RunAsync(service, "session.open", new { filePath = path });

        Assert.False(response.Success);
        Assert.Equal(nameof(ArgumentException), response.ErrorType);
    }

    [Fact]
    public async Task Open_ReportsAMissingFileWithAPointerToCreate()
    {
        using var service = new WordMcpService();
        var missing = Path.Combine(Path.GetTempPath(), $"wordmcp-{Guid.NewGuid():N}.docx");

        var response = await RunAsync(service, "session.open", new { filePath = missing });

        Assert.False(response.Success);
        Assert.Equal(nameof(FileNotFoundException), response.ErrorType);
        Assert.Contains("session.create", response.ErrorMessage!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Create_RefusesToOverwriteAnExistingFile()
    {
        using var service = new WordMcpService();
        var path = Path.Combine(Path.GetTempPath(), $"wordmcp-{Guid.NewGuid():N}.docx");
        await File.WriteAllTextAsync(path, "not really a document");

        try
        {
            var response = await RunAsync(service, "session.create", new { filePath = path });

            Assert.False(response.Success);
            Assert.Equal(nameof(IOException), response.ErrorType);
            Assert.Contains("session.open", response.ErrorMessage!, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("session.save")]
    [InlineData("session.close")]
    public async Task SessionScopedCommandsRequireASessionId(string command)
    {
        using var service = new WordMcpService();

        var response = await RunAsync(service, command);

        Assert.False(response.Success);
        Assert.Equal(nameof(ArgumentException), response.ErrorType);
    }

    [Fact]
    public async Task Save_ReportsAnUnknownSessionId()
    {
        using var service = new WordMcpService();

        var response = await RunAsync(service, "session.save", sessionId: "word-does-not-exist");

        Assert.False(response.Success);
        Assert.Equal(nameof(KeyNotFoundException), response.ErrorType);
    }

    [Fact]
    public async Task Close_ReportsAnUnknownSessionIdAsAnUnsuccessfulResultRatherThanAnError()
    {
        using var service = new WordMcpService();

        var payload = JsonDocument.Parse(
            (await RunAsync(service, "session.close", sessionId: "word-does-not-exist")).Result!).RootElement;

        Assert.False(payload.GetProperty("success").GetBoolean());
        Assert.Contains("word-does-not-exist", payload.GetProperty("message").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Test_ReportsAMissingFileAsNotOpenable()
    {
        using var service = new WordMcpService();
        var missing = Path.Combine(Path.GetTempPath(), $"wordmcp-{Guid.NewGuid():N}.docx");

        var payload = Payload(await RunAsync(service, "session.test", new { filePath = missing }));

        Assert.False(payload.GetProperty("canOpen").GetBoolean());
        Assert.False(payload.GetProperty("exists").GetBoolean());
    }

    [Fact]
    public async Task Test_ReportsAnUnsupportedExtension()
    {
        using var service = new WordMcpService();
        var path = Path.Combine(Path.GetTempPath(), $"wordmcp-{Guid.NewGuid():N}.txt");
        await File.WriteAllTextAsync(path, "plain text");

        try
        {
            var payload = Payload(await RunAsync(service, "session.test", new { filePath = path }));

            Assert.False(payload.GetProperty("canOpen").GetBoolean());
            Assert.Contains(
                payload.GetProperty("problems").EnumerateArray(),
                p => p.GetString()!.Contains(".txt", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void GetBatch_RequiresASessionId()
    {
        using var service = new WordMcpService();

        Assert.Throws<ArgumentException>(() => service.GetBatch(null));
    }

    [Fact]
    public void GetBatch_ReportsAnUnknownSessionId()
    {
        using var service = new WordMcpService();

        Assert.Throws<KeyNotFoundException>(() => service.GetBatch("word-does-not-exist"));
    }

    [Fact]
    public void AFreshServiceIsNotYetIdle()
    {
        var clock = new ManualClock(DateTimeOffset.UnixEpoch);
        using var service = new WordMcpService(idleTimeout: TimeSpan.FromMinutes(5), clock: clock);

        Assert.False(service.IsIdle);
    }

    [Fact]
    public void AServiceBecomesIdleOnceTheTimeoutElapsesWithoutSessions()
    {
        var clock = new ManualClock(DateTimeOffset.UnixEpoch);
        using var service = new WordMcpService(idleTimeout: TimeSpan.FromMinutes(5), clock: clock);

        clock.Advance(TimeSpan.FromMinutes(5));

        Assert.True(service.IsIdle);
    }

    [Fact]
    public async Task ARequestPostponesTheIdleDeadline()
    {
        var clock = new ManualClock(DateTimeOffset.UnixEpoch);
        using var service = new WordMcpService(idleTimeout: TimeSpan.FromMinutes(5), clock: clock);

        clock.Advance(TimeSpan.FromMinutes(4));
        await RunAsync(service, "service.ping");
        clock.Advance(TimeSpan.FromMinutes(4));

        Assert.False(service.IsIdle);
    }

    [Fact]
    public async Task ProcessAsync_RejectsCallsAfterDispose()
    {
        var service = new WordMcpService();
        service.Dispose();

        await Assert.ThrowsAsync<ObjectDisposedException>(
            () => service.ProcessAsync(new ServiceRequest { Command = "service.ping" }));
    }

    [Fact]
    public void DisposeIsIdempotent()
    {
        var service = new WordMcpService();

        service.Dispose();
        service.Dispose();
    }

    private sealed class ManualClock(DateTimeOffset now) : TimeProvider
    {
        private DateTimeOffset _now = now;

        public override DateTimeOffset GetUtcNow() => _now;

        public void Advance(TimeSpan delta) => _now += delta;
    }
}

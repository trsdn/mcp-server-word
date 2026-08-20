using System.Text.Json;
using Xunit;

namespace WordMcp.Service.Tests;

/// <summary>
/// Session lifecycle against a real Word instance. These check the promise the whole service
/// exists for: a document opened once stays open and is handed back to the next caller.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class SessionLifecycleIntegrationTests : IDisposable
{
    private readonly WordMcpService _service = new();
    private readonly string _directory;

    public SessionLifecycleIntegrationTests()
        => _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-service-" + Guid.NewGuid().ToString("N")))
            .FullName;

    public void Dispose()
    {
        _service.Dispose();

        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
            // A stray Word lock file must not fail the test that already passed.
        }
    }

    private string Path_(string name) => System.IO.Path.Combine(_directory, name);

    private async Task<JsonElement> RunAsync(string command, object? args = null, string? sessionId = null)
    {
        var response = await _service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args is null ? null : ServiceProtocol.Serialize(args)
        });

        Assert.True(response.Success, response.ErrorMessage);
        Assert.NotNull(response.Result);
        return JsonDocument.Parse(response.Result).RootElement.Clone();
    }

    [Fact]
    public async Task Create_OpensASessionForANewDocument()
    {
        var path = Path_("created.docx");

        var payload = await RunAsync("session.create", new { filePath = path });

        Assert.True(payload.GetProperty("success").GetBoolean());
        Assert.StartsWith("word-", payload.GetProperty("sessionId").GetString()!, StringComparison.Ordinal);
        Assert.Equal(path, payload.GetProperty("filePath").GetString());
        Assert.True(File.Exists(path));
    }

    [Fact]
    public async Task List_ShowsAnOpenSession()
    {
        var sessionId = (await RunAsync("session.create", new { filePath = Path_("listed.docx") }))
            .GetProperty("sessionId").GetString();

        var payload = await RunAsync("session.list");

        Assert.Equal(1, payload.GetProperty("count").GetInt32());
        Assert.Equal(sessionId, payload.GetProperty("sessions")[0].GetProperty("sessionId").GetString());
    }

    [Fact]
    public async Task Open_ReusesTheSessionOfAnAlreadyOpenDocument()
    {
        var path = Path_("reused.docx");
        var first = (await RunAsync("session.create", new { filePath = path }))
            .GetProperty("sessionId").GetString();

        var second = await RunAsync("session.open", new { filePath = path });

        Assert.True(second.GetProperty("reused").GetBoolean());
        Assert.Equal(first, second.GetProperty("sessionId").GetString());
        Assert.Equal(1, (await RunAsync("session.list")).GetProperty("count").GetInt32());
    }

    [Fact]
    public async Task ADocumentStaysOpenAcrossRequests()
    {
        var sessionId = (await RunAsync("session.create", new { filePath = Path_("persistent.docx") }))
            .GetProperty("sessionId").GetString();

        // The point of the service: the batch behind the id survives, so a later request finds
        // the same live Word instance instead of reopening the file.
        var batch = _service.GetBatch(sessionId);
        Assert.True(batch.IsWordProcessAlive());

        await RunAsync("service.ping");

        Assert.Same(batch, _service.GetBatch(sessionId));
    }

    [Fact]
    public async Task Save_PersistsTheDocumentWithoutClosingTheSession()
    {
        var path = Path_("saved.docx");
        var sessionId = (await RunAsync("session.create", new { filePath = path }))
            .GetProperty("sessionId").GetString();

        var payload = await RunAsync("session.save", sessionId: sessionId);

        Assert.True(payload.GetProperty("success").GetBoolean());
        Assert.True(new FileInfo(path).Length > 0);
        Assert.Equal(1, (await RunAsync("session.list")).GetProperty("count").GetInt32());
    }

    [Fact]
    public async Task Close_RemovesTheSession()
    {
        var sessionId = (await RunAsync("session.create", new { filePath = Path_("closed.docx") }))
            .GetProperty("sessionId").GetString();

        var payload = await RunAsync("session.close", new { save = true }, sessionId);

        Assert.True(payload.GetProperty("success").GetBoolean());
        Assert.True(payload.GetProperty("saved").GetBoolean());
        Assert.Equal(0, (await RunAsync("session.list")).GetProperty("count").GetInt32());
        Assert.Throws<KeyNotFoundException>(() => _service.GetBatch(sessionId));
    }

    [Fact]
    public async Task AnOpenSessionKeepsTheServiceFromGoingIdle()
    {
        await RunAsync("session.create", new { filePath = Path_("busy.docx") });

        using var busy = new WordMcpService(idleTimeout: TimeSpan.Zero);

        // A zero timeout makes a service with no sessions idle immediately; this one has one.
        Assert.True(busy.IsIdle);
        Assert.False(_service.IsIdle);
        Assert.Equal(1, _service.SessionCount);
    }

    [Fact]
    public async Task Test_ConfirmsThatAClosedDocumentCanBeOpened()
    {
        var path = Path_("testable.docx");
        var sessionId = (await RunAsync("session.create", new { filePath = path }))
            .GetProperty("sessionId").GetString();
        await RunAsync("session.close", new { save = true }, sessionId);

        var payload = await RunAsync("session.test", new { filePath = path });

        Assert.True(payload.GetProperty("canOpen").GetBoolean());
        Assert.True(payload.GetProperty("exists").GetBoolean());
        Assert.Empty(payload.GetProperty("problems").EnumerateArray());
    }

    [Fact]
    public async Task Dispose_SavesAndClosesEveryOpenSession()
    {
        var path = Path_("shutdown.docx");
        var service = new WordMcpService();
        await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = ServiceProtocol.Serialize(new { filePath = path })
        });

        service.Dispose();

        // A saved and released document can be opened again; a leaked one would still be locked.
        var reopened = await RunAsync("session.test", new { filePath = path });
        Assert.True(reopened.GetProperty("canOpen").GetBoolean());
    }
}

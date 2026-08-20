using System.Text.Json;
using WordMcp.McpServer.Tools;
using Xunit;

namespace WordMcp.McpServer.Tests;

/// <summary>
/// The one path the bridge tests cannot fake: a document really being created, saved, listed and
/// closed through the file tool. Requires Word.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class WordFileToolIntegrationTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), $"wordmcp-tool-{Guid.NewGuid():N}");

    public WordFileToolIntegrationTests() => Directory.CreateDirectory(_directory);

    public void Dispose()
    {
        ServiceBridge.Shutdown();

        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
            // Word may still be releasing the file; the temp directory is disposable anyway.
        }
    }

    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement.Clone();

    [Fact]
    public void ADocumentSurvivesTheWholeRoundTrip()
    {
        var path = Path.Combine(_directory, "report.docx");

        var created = Parse(WordFileTool.File(WordFileAction.Create, path));
        Assert.True(created.GetProperty("success").GetBoolean(), created.ToString());
        var sessionId = created.GetProperty("sessionId").GetString()!;

        try
        {
            var listed = Parse(WordFileTool.File(WordFileAction.List));
            Assert.Equal(1, listed.GetProperty("count").GetInt32());
            Assert.Equal(sessionId, listed.GetProperty("sessions")[0].GetProperty("sessionId").GetString());

            Assert.True(Parse(WordFileTool.File(WordFileAction.Save, session_id: sessionId))
                .GetProperty("success").GetBoolean());

            // Opening the same path again must hand back the session that is already there rather
            // than starting a second Word on a locked file.
            var reopened = Parse(WordFileTool.File(WordFileAction.Open, path));
            Assert.True(reopened.GetProperty("reused").GetBoolean());
            Assert.Equal(sessionId, reopened.GetProperty("sessionId").GetString());
        }
        finally
        {
            Assert.True(Parse(WordFileTool.File(WordFileAction.Close, session_id: sessionId, save: true))
                .GetProperty("success").GetBoolean());
        }

        Assert.True(File.Exists(path));
        Assert.Equal(0, Parse(WordFileTool.File(WordFileAction.List)).GetProperty("count").GetInt32());

        var tested = Parse(WordFileTool.File(WordFileAction.Test, path));
        Assert.True(tested.GetProperty("canOpen").GetBoolean(), tested.ToString());
    }
}

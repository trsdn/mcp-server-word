using System.Text.Json;
using WordMcp.McpServer.Tools;
using Xunit;

namespace WordMcp.McpServer.Tests;

/// <summary>
/// The generated tools no longer touch a COM object: they pack their arguments, hand them to the
/// bridge and read a result back. These tests are what proves that packing and unpacking agree,
/// because nothing else does — a mismatched name or a dropped default only shows up when a real
/// document comes back changed. Requires Word.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class GeneratedToolIntegrationTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), $"wordmcp-generated-{Guid.NewGuid():N}");

    public GeneratedToolIntegrationTests() => Directory.CreateDirectory(_directory);

    public void Dispose()
    {
        ServiceBridge.Shutdown();
        Environment.SetEnvironmentVariable(ServiceBridge.ModeVariable, null);

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
    public void TextWrittenThroughAToolComesBackThroughAnother()
    {
        var path = Path.Combine(_directory, "generated.docx");
        var created = Parse(WordFileTool.File(WordFileAction.Create, path));
        Assert.True(created.GetProperty("success").GetBoolean(), created.ToString());
        var session = created.GetProperty("sessionId").GetString()!;

        try
        {
            var appended = Parse(WordTextTool.Text(WordTextAction.Append, session, text: "Guten Morgen."));
            Assert.True(appended.GetProperty("success").GetBoolean(), appended.ToString());

            var read = Parse(WordTextTool.Text(WordTextAction.Get, session));
            Assert.Contains("Guten Morgen.", read.GetProperty("text").GetString(), StringComparison.Ordinal);

            // A replace exercises the arguments that carry a default on both sides of the seam.
            var replaced = Parse(WordTextTool.Text(
                WordTextAction.Replace, session, search_text: "Morgen", replace_text: "Abend"));
            Assert.True(replaced.GetProperty("success").GetBoolean(), replaced.ToString());

            var after = Parse(WordTextTool.Text(WordTextAction.Get, session));
            Assert.Contains("Guten Abend.", after.GetProperty("text").GetString(), StringComparison.Ordinal);
        }
        finally
        {
            WordFileTool.File(WordFileAction.Close, session_id: session, save: false);
        }
    }

    [Fact]
    public void AMissingArgumentIsStillTheCallersParameterName()
    {
        var path = Path.Combine(_directory, "missing.docx");
        var session = Parse(WordFileTool.File(WordFileAction.Create, path)).GetProperty("sessionId").GetString()!;

        try
        {
            // add needs a name; the message has to name the parameter the caller typed, not
            // whatever the service calls it internally.
            var failed = Parse(WordBookmarkTool.Bookmark(WordBookmarkAction.Add, session, paragraph_index: 1));

            Assert.False(failed.GetProperty("success").GetBoolean());
            Assert.Equal("ArgumentException", failed.GetProperty("errorType").GetString());
            Assert.Contains("name is required", failed.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
        }
        finally
        {
            WordFileTool.File(WordFileAction.Close, session_id: session, save: false);
        }
    }

    [Fact]
    public void TheSameToolWorksThroughTheDaemon()
    {
        Environment.SetEnvironmentVariable(ServiceBridge.ModeVariable, "daemon");
        var path = Path.Combine(_directory, "daemon.docx");

        var created = Parse(WordFileTool.File(WordFileAction.Create, path));
        Assert.True(created.GetProperty("success").GetBoolean(), created.ToString());
        var session = created.GetProperty("sessionId").GetString()!;

        try
        {
            var appended = Parse(WordTextTool.Text(WordTextAction.Append, session, text: "Aus dem Dienst."));
            Assert.True(appended.GetProperty("success").GetBoolean(), appended.ToString());

            var read = Parse(WordTextTool.Text(WordTextAction.Get, session));
            Assert.Contains("Aus dem Dienst.", read.GetProperty("text").GetString(), StringComparison.Ordinal);
        }
        finally
        {
            WordFileTool.File(WordFileAction.Close, session_id: session, save: false);
            ServiceBridge.Invoke("service.shutdown");
        }
    }
}

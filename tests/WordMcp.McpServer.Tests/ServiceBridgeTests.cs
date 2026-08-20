using System.Text.Json;
using WordMcp.McpServer.Tools;
using Xunit;

namespace WordMcp.McpServer.Tests;

/// <summary>
/// The bridge and the file tool that now runs on it. None of these open Word: every case here is
/// answered before a document is ever touched, which is exactly the part that used to be
/// duplicated between the tool and the service.
/// </summary>
public class ServiceBridgeTests
{
    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement.Clone();

    private static JsonElement File(WordFileAction action, string? path = null, string? sessionId = null, bool save = false)
        => Parse(WordFileTool.File(action, path, sessionId, save));

    [Fact]
    public void TheBridgeAnswersAServiceCommand()
    {
        var payload = (JsonElement)ServiceBridge.Invoke("service.ping");

        Assert.True(payload.GetProperty("pong").GetBoolean());
        Assert.Equal(Environment.ProcessId, payload.GetProperty("processId").GetInt32());
    }

    [Fact]
    public void AnUnknownCommandBecomesAnException()
    {
        var ex = Assert.Throws<ServiceCommandException>(() => ServiceBridge.Invoke("session.frobnicate"));

        Assert.Contains("session.frobnicate", ex.Message, StringComparison.Ordinal);
        Assert.Equal("NotSupportedException", ex.ErrorType);
    }

    [Fact]
    public void TheReportedErrorTypeSurvivesTheTrip()
    {
        var json = Parse(WordToolsBase.Execute(
            "file", "open", () => throw new ServiceCommandException("gone", "FileNotFoundException")));

        // Without this the caller would see the name of the wrapper instead of the real failure.
        Assert.Equal("FileNotFoundException", json.GetProperty("errorType").GetString());
        Assert.Equal("gone", json.GetProperty("errorMessage").GetString());
    }

    [Fact]
    public void AWrapperWithoutATypeFallsBackToItsOwnName()
    {
        var json = Parse(WordToolsBase.Execute("file", "open", () => throw new ServiceCommandException("gone")));

        Assert.Equal("ServiceCommandException", json.GetProperty("errorType").GetString());
    }

    [Fact]
    public void ListingSessionsGoesThroughTheService()
    {
        var json = File(WordFileAction.List);

        Assert.True(json.GetProperty("success").GetBoolean());
        Assert.Equal(0, json.GetProperty("count").GetInt32());
        Assert.Empty(json.GetProperty("sessions").EnumerateArray());
    }

    [Theory]
    [InlineData(WordFileAction.Open)]
    [InlineData(WordFileAction.Create)]
    public void ARelativePathIsRejectedBeforeTheServiceSeesIt(WordFileAction action)
    {
        var json = File(action, @"docs\report.docx");

        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.Equal("ArgumentException", json.GetProperty("errorType").GetString());
        Assert.Contains("absolute", json.GetProperty("errorMessage").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnsupportedExtensionIsRejectedForCreate()
    {
        var json = File(WordFileAction.Create, @"C:\docs\deck.pptx");

        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.Contains("Unsupported file type", json.GetProperty("errorMessage").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public void OpeningAMissingDocumentReportsItAsMissing()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");

        var json = File(WordFileAction.Open, missing);

        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.Equal("FileNotFoundException", json.GetProperty("errorType").GetString());
    }

    [Fact]
    public void CreatingOverAnExistingFileIsRefused()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        System.IO.File.WriteAllText(path, "not really a document");

        try
        {
            var json = File(WordFileAction.Create, path);

            Assert.False(json.GetProperty("success").GetBoolean());
            Assert.Equal("IOException", json.GetProperty("errorType").GetString());
            Assert.Contains("already exists", json.GetProperty("errorMessage").GetString()!, StringComparison.Ordinal);
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Theory]
    [InlineData(WordFileAction.Save)]
    [InlineData(WordFileAction.Close)]
    public void ASessionScopedActionNeedsASessionId(WordFileAction action)
    {
        var json = File(action);

        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.Equal("ArgumentException", json.GetProperty("errorType").GetString());
        Assert.Contains("session_id", json.GetProperty("errorMessage").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public void ClosingAnUnknownSessionIsReportedNotThrown()
    {
        var json = File(WordFileAction.Close, sessionId: "word-doesnotexist");

        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.Contains("word-doesnotexist", json.GetProperty("message").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public void SavingAnUnknownSessionIsAnError()
    {
        var json = File(WordFileAction.Save, sessionId: "word-doesnotexist");

        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.Equal("KeyNotFoundException", json.GetProperty("errorType").GetString());
    }

    [Fact]
    public void TestingAMissingFileListsTheProblem()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");

        var json = File(WordFileAction.Test, missing);

        Assert.False(json.GetProperty("canOpen").GetBoolean());
        Assert.False(json.GetProperty("exists").GetBoolean());
        Assert.Contains(
            json.GetProperty("problems").EnumerateArray(),
            p => p.GetString()!.Contains("does not exist", StringComparison.Ordinal));
    }

    [Fact]
    public void TestingReportsAnUnsupportedExtensionAsAFindingRatherThanAFailure()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.pptx");
        System.IO.File.WriteAllText(path, "deck");

        try
        {
            var json = File(WordFileAction.Test, path);

            // Answering "what would happen" is the point, so this must not throw.
            Assert.False(json.GetProperty("canOpen").GetBoolean());
            Assert.True(json.GetProperty("exists").GetBoolean());
            Assert.Contains(
                json.GetProperty("problems").EnumerateArray(),
                p => p.GetString()!.Contains("Unsupported extension", StringComparison.Ordinal));
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void TestingNeedsAPath()
    {
        var json = File(WordFileAction.Test);

        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.Equal("ArgumentException", json.GetProperty("errorType").GetString());
    }

    [Fact]
    public void ABatchStillNeedsASessionId()
    {
        var ex = Assert.Throws<ArgumentException>(() => WordToolsBase.Batch(null));

        Assert.Contains("session_id", ex.Message, StringComparison.Ordinal);
    }
}

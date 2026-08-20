using System.Text.Json;
using WordMcp.McpServer.Tools;
using Xunit;

namespace WordMcp.McpServer.Tests;

public class WordToolsBaseTests
{
    [Fact]
    public void Execute_SerializesSuccessfulResults()
    {
        var json = WordToolsBase.Execute("text", "get", () => new { success = true, text = "hello" });

        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("hello", doc.RootElement.GetProperty("text").GetString());
    }

    [Fact]
    public void Execute_ConvertsExceptionsIntoStructuredErrors()
    {
        var json = WordToolsBase.Execute("table", "read", () => throw new InvalidOperationException("boom"));

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.True(root.GetProperty("isError").GetBoolean());
        Assert.Equal("table", root.GetProperty("tool").GetString());
        Assert.Equal("read", root.GetProperty("action").GetString());
        Assert.Equal("InvalidOperationException", root.GetProperty("errorType").GetString());
        Assert.Equal("boom", root.GetProperty("errorMessage").GetString());
    }

    [Fact]
    public void Execute_AddsGuidanceForTimeouts()
    {
        var json = WordToolsBase.Execute("document", "get-info", () => throw new TimeoutException("Operation timed out."));

        using var doc = JsonDocument.Parse(json);
        Assert.Contains("dialog", doc.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Serialize_UsesCamelCaseAndSkipsNulls()
    {
        var json = WordToolsBase.Serialize(new { filePath = @"C:\a.docx", message = (string?)null });

        Assert.Contains("\"filePath\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("message", json, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Batch_RequiresSessionId(string? sessionId)
    {
        var ex = Assert.Throws<ArgumentException>(() => WordToolsBase.Batch(sessionId));
        Assert.Contains("session_id", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Batch_ReportsUnknownSessionId()
        => Assert.Throws<KeyNotFoundException>(() => WordToolsBase.Batch("word-doesnotexist"));

    [Theory]
    [InlineData("report.docx")]
    [InlineData(@"..\report.docx")]
    [InlineData(@"docs\report.docx")]
    public void ValidatePath_RejectsRelativePaths(string path)
    {
        var ex = Assert.Throws<ArgumentException>(() => WordToolsBase.ValidatePath(path, mustExist: false));
        Assert.Contains("absolute", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ValidatePath_RejectsEmptyPath()
        => Assert.Throws<ArgumentException>(() => WordToolsBase.ValidatePath("  ", mustExist: false));

    [Theory]
    [InlineData(@"C:\docs\slides.pptx")]
    [InlineData(@"C:\docs\notes.txt")]
    [InlineData(@"C:\docs\report.pdf")]
    public void ValidatePath_RejectsUnsupportedExtensions(string path)
    {
        var ex = Assert.Throws<ArgumentException>(() => WordToolsBase.ValidatePath(path, mustExist: false));
        Assert.Contains("Unsupported file type", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"C:\docs\report.docx")]
    [InlineData(@"C:\docs\macro.docm")]
    [InlineData(@"C:\docs\legacy.doc")]
    [InlineData(@"C:\docs\template.dotx")]
    public void ValidatePath_AcceptsSupportedExtensions(string path)
        => Assert.Equal(Path.GetFullPath(path), WordToolsBase.ValidatePath(path, mustExist: false));

    [Fact]
    public void ValidatePath_NormalizesTraversalSegments()
    {
        var result = WordToolsBase.ValidatePath(@"C:\docs\sub\..\report.docx", mustExist: false);
        Assert.Equal(@"C:\docs\report.docx", result);
    }

    [Fact]
    public void ValidatePath_ThrowsWhenExistenceIsRequired()
    {
        var missing = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        Assert.Throws<FileNotFoundException>(() => WordToolsBase.ValidatePath(missing, mustExist: true));
    }
}

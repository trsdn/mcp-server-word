using System.Text.Json;
using WordMcp.Core.Models;
using Xunit;

namespace WordMcp.Core.Tests;

public class ResultModelTests
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    [Fact]
    public void OperationResult_DefaultsToSuccess()
    {
        var result = new OperationResult();

        Assert.True(result.Success);
        Assert.Null(result.Message);
    }

    [Fact]
    public void Results_SerializeAsCamelCase()
    {
        var json = JsonSerializer.Serialize(
            new DocumentInfoResult { FilePath = @"C:\a.docx", WordCount = 12 }, Options);

        Assert.Contains("\"filePath\"", json, StringComparison.Ordinal);
        Assert.Contains("\"wordCount\":12", json, StringComparison.Ordinal);
    }

    [Fact]
    public void Results_OmitNullMessage()
    {
        var json = JsonSerializer.Serialize(new OperationResult(), Options);

        Assert.DoesNotContain("message", json, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"success\":true", json, StringComparison.Ordinal);
    }

    [Fact]
    public void CollectionResults_DefaultToEmptyNotNull()
    {
        Assert.Empty(new ParagraphListResult().Paragraphs);
        Assert.Empty(new TableListResult().Tables);
        Assert.Empty(new TableReadResult().Rows);
        Assert.Empty(new FindResult().Matches);
    }

    [Fact]
    public void TableReadResult_SerializesRowsAsNestedArrays()
    {
        var result = new TableReadResult
        {
            Table = new TableInfo { Index = 1, RowCount = 1, ColumnCount = 2, Style = "Table Grid" },
            Rows = [["a", "b"]]
        };

        var json = JsonSerializer.Serialize(result, Options);

        Assert.Contains("\"rows\":[[\"a\",\"b\"]]", json, StringComparison.Ordinal);
    }
}

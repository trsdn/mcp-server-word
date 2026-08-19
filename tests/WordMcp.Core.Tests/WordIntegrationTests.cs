using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Document;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Table;
using WordMcp.Core.Commands.Text;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// End-to-end tests against a real Word installation. Excluded from CI via the
/// <c>RequiresWord</c> category because GitHub runners have no Office.
/// </summary>
[Collection(WordCollection.Name)]
[Trait("Category", "RequiresWord")]
public class WordIntegrationTests(WordDocumentFixture fixture)
{
    private static readonly DocumentCommands Documents = new();
    private static readonly TextCommands Texts = new();
    private static readonly ParagraphCommands Paragraphs = new();
    private static readonly TableCommands Tables = new();

    [Fact]
    public void Document_IsCreatedOnDisk()
    {
        fixture.Sessions.Save(fixture.SessionId);
        Assert.True(File.Exists(fixture.FilePath));
    }

    [Fact]
    public void GetInfo_ReturnsStatistics()
    {
        // Exercises the dynamic ComputeStatistics path over WdStatistic.
        var info = Documents.GetInfo(fixture.Batch);

        Assert.True(info.Success);
        Assert.Equal(fixture.FilePath, info.FilePath, ignoreCase: true);
        Assert.Equal("integration.docx", info.FileName);
        Assert.True(info.PageCount >= 1);
        Assert.True(info.ParagraphCount >= 1);
        Assert.True(info.CharacterCount >= 0);
    }

    [Fact]
    public void Properties_RoundTrip()
    {
        // Exercises the BuiltInDocumentProperties indexer in both directions.
        Documents.SetProperties(fixture.Batch, title: "Integration Title", author: "WordMcp", subject: "Testing");

        var properties = Documents.GetProperties(fixture.Batch);

        Assert.Equal("Integration Title", properties.Title);
        Assert.Equal("WordMcp", properties.Author);
        Assert.Equal("Testing", properties.Subject);
    }

    [Fact]
    public void Text_AppendFindReplaceAndFormat()
    {
        Texts.Append(fixture.Batch, "The quick brown fox jumps over the lazy dog.");

        var found = Texts.Find(fixture.Batch, "brown");
        Assert.Equal(1, found.MatchCount);
        var match = Assert.Single(found.Matches);
        Assert.True(match.Start >= 0);
        Assert.True(match.End > match.Start);
        Assert.Contains("brown", match.Context, StringComparison.Ordinal);

        // Exercises Range.Find.Execute with positional arguments.
        var replaced = Texts.Replace(fixture.Batch, "brown", "red");
        Assert.Equal(1, replaced.ReplacementCount);
        Assert.Contains("red fox", Texts.Get(fixture.Batch).Text, StringComparison.Ordinal);

        var formatted = Texts.Format(fixture.Batch, match.Start, match.Start + 3, bold: true, color: "#0078D4");
        Assert.True(formatted.Success);
    }

    [Fact]
    public void Text_GetHonoursMaxLength()
    {
        Texts.Append(fixture.Batch, new string('x', 200));

        var text = Texts.Get(fixture.Batch, maxLength: 25);

        Assert.True(text.Text.Length <= 25);
        Assert.True(text.Truncated);
    }

    [Fact]
    public void Paragraph_AddStyleAlignmentAndDelete()
    {
        // Exercises the Range.Style setter/getter pair, which does not work via get_Style().
        var added = Paragraphs.Add(fixture.Batch, "Integration Heading", style: "Heading 1", alignment: "center");

        Assert.True(added.Success);
        Assert.NotNull(added.Paragraph);
        Assert.Equal("Integration Heading", added.Paragraph!.Text);
        // Word reports built-in styles under their localized name ("Überschrift 1" on a German
        // installation), so compare against a paragraph that carries the default style instead of
        // asserting a specific name.
        var plain = Paragraphs.Add(fixture.Batch, "Integration Plain");
        Assert.NotEqual(plain.Paragraph!.Style, added.Paragraph.Style);
        Paragraphs.Delete(fixture.Batch, plain.Paragraph.Index);
        Assert.Equal("center", added.Paragraph.Alignment);

        var index = added.Paragraph.Index;

        Paragraphs.SetAlignment(fixture.Batch, index, "right");
        var list = Paragraphs.List(fixture.Batch);
        Assert.Equal("right", list.Paragraphs.Single(p => p.Index == index).Alignment);

        var before = Paragraphs.List(fixture.Batch).TotalCount;
        Paragraphs.Delete(fixture.Batch, index);
        Assert.Equal(before - 1, Paragraphs.List(fixture.Batch).TotalCount);
    }

    [Fact]
    public void Table_CreateWriteReadAndGrow()
    {
        // Exercises Tables.Add(range, rows, cols) and the end-of-cell marker cleanup.
        var created = Tables.Create(fixture.Batch, rows: 2, columns: 3, style: "Table Grid");

        Assert.True(created.Success);
        Assert.NotNull(created.Table);
        var index = created.Table!.Index;
        Assert.Equal(2, created.Table.RowCount);
        Assert.Equal(3, created.Table.ColumnCount);

        Tables.SetCell(fixture.Batch, index, row: 1, column: 1, text: "Header");
        Tables.SetCell(fixture.Batch, index, row: 2, column: 3, text: "Corner");

        var read = Tables.Read(fixture.Batch, index);
        Assert.Equal(2, read.Rows.Count);
        Assert.Equal("Header", read.Rows[0][0]);
        Assert.Equal("Corner", read.Rows[1][2]);
        // Cell text must not contain the \r\a end-of-cell marker.
        Assert.DoesNotContain(read.Rows[0][0], "\a", StringComparison.Ordinal);

        var grown = Tables.AddRow(fixture.Batch, index, ["a", "b", "c"]);
        Assert.Equal(3, grown.Table!.RowCount);
        Assert.Equal("b", Tables.Read(fixture.Batch, index).Rows[2][1]);

        var shrunk = Tables.DeleteRow(fixture.Batch, index, row: 3);
        Assert.Equal(2, shrunk.Table!.RowCount);

        Assert.Contains(Tables.List(fixture.Batch).Tables, t => t.Index == index);
    }

    [Fact]
    public void ExportPdf_WritesFileAndLeavesDocumentOpen()
    {
        // Exercises ExportAsFixedFormat, which must not change the open document's path.
        var target = Path.Combine(fixture.DirectoryPath, "export.pdf");

        var result = Documents.ExportPdf(fixture.Batch, target);

        Assert.True(result.Success);
        Assert.True(File.Exists(target));
        Assert.True(result.FileSizeBytes > 0);
        Assert.Equal(fixture.FilePath, Documents.GetInfo(fixture.Batch).FilePath, ignoreCase: true);
    }

    [Fact]
    public void SaveAs_WritesCopyAndKeepsOriginalPath()
    {
        // SaveAs2 round-trip: the open document must end up pointing at the original file again.
        var target = Path.Combine(fixture.DirectoryPath, "copy.rtf");

        var result = Documents.SaveAs(fixture.Batch, target);

        Assert.True(result.Success);
        Assert.True(File.Exists(target));
        Assert.Equal(fixture.FilePath, Documents.GetInfo(fixture.Batch).FilePath, ignoreCase: true);
    }
}

/// <summary>
/// Session lifecycle tests that need their own Word instances.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class WordSessionIntegrationTests : IDisposable
{
    private readonly string _directory =
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-it-" + Guid.NewGuid().ToString("N"))).FullName;

    private readonly SessionManager _sessions = new();

    public void Dispose()
    {
        _sessions.Dispose();
        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
            // Best effort cleanup.
        }
    }

    [Fact]
    public void CreateSaveCloseAndReopen()
    {
        var path = Path.Combine(_directory, "lifecycle.docx");

        var created = _sessions.Create(path);
        Assert.StartsWith("word-", created.SessionId, StringComparison.Ordinal);

        new TextCommands().Append(_sessions.GetBatch(created.SessionId), "Persisted content");
        Assert.True(_sessions.Close(created.SessionId, save: true));
        Assert.True(File.Exists(path));
        Assert.Empty(_sessions.List());

        var reopened = _sessions.Open(path);
        Assert.Contains("Persisted content",
            new TextCommands().Get(_sessions.GetBatch(reopened.SessionId)).Text, StringComparison.Ordinal);

        Assert.Equal(reopened.SessionId, _sessions.FindByPath(path)?.SessionId);
        Assert.Single(_sessions.List());

        Assert.True(_sessions.Close(reopened.SessionId, save: false));
    }

    [Fact]
    public void GetBatch_ThrowsForClosedSession()
    {
        var path = Path.Combine(_directory, "closed.docx");
        var session = _sessions.Create(path);

        _sessions.Close(session.SessionId, save: false);

        Assert.Throws<KeyNotFoundException>(() => _sessions.GetBatch(session.SessionId));
    }
}

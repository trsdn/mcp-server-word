using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.ContentControl;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Text;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the content control commands against a real Word instance.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class ContentControlIntegrationTests : IDisposable
{
    private static readonly ContentControlCommands Controls = new();
    private static readonly ParagraphCommands Paragraphs = new();
    private static readonly TextCommands Texts = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public ContentControlIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-cc-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "controls.docx")).SessionId;
    }

    [Fact]
    public void ContentControl_AddListSetTextAndDeleteRoundTrip()
    {
        int index = Paragraphs.Add(Batch, "Placeholder for the customer.").Paragraph!.Index;

        var added = Controls.Add(Batch, index, "rich-text", "CustomerName", "Customer");

        Assert.True(added.Success);
        Assert.Equal(1, added.AffectedCount);
        Assert.Equal("CustomerName", added.Controls[0].Tag);
        Assert.Equal("Customer", added.Controls[0].Title);
        Assert.Equal("rich-text", added.Controls[0].Type);
        Assert.NotEmpty(added.Controls[0].Id);

        var list = Controls.List(Batch);
        Assert.Equal(1, list.TotalCount);
        Assert.Equal("CustomerName", list.Controls[0].Tag);

        var filled = Controls.SetText(Batch, "Contoso Ltd", "CustomerName");
        Assert.Equal(1, filled.AffectedCount);
        Assert.Equal("Contoso Ltd", filled.Controls[0].Text);

        Assert.Equal("Contoso Ltd", Controls.Get(Batch, "CustomerName").Controls[0].Text);
        Assert.Contains("Contoso Ltd", Texts.Get(Batch).Text, StringComparison.Ordinal);

        var deleted = Controls.Delete(Batch, "CustomerName");
        Assert.Equal(1, deleted.AffectedCount);
        Assert.Empty(Controls.List(Batch).Controls);

        // Deleting the control alone must leave the text behind.
        Assert.Contains("Contoso Ltd", Texts.Get(Batch).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentControl_SetTextByTagFillsEveryMatch()
    {
        int first = Paragraphs.Add(Batch, "Header name.").Paragraph!.Index;
        int second = Paragraphs.Add(Batch, "Signature name.").Paragraph!.Index;

        Controls.Add(Batch, first, "rich-text", "CustomerName");
        Controls.Add(Batch, second, "rich-text", "CustomerName");

        var filled = Controls.SetText(Batch, "Contoso Ltd", "CustomerName");

        Assert.Equal(2, filled.AffectedCount);
        Assert.All(filled.Controls, c => Assert.Equal("Contoso Ltd", c.Text));
    }

    [Fact]
    public void ContentControl_SetTextByIdHitsExactlyOne()
    {
        int first = Paragraphs.Add(Batch, "Header name.").Paragraph!.Index;
        int second = Paragraphs.Add(Batch, "Signature name.").Paragraph!.Index;

        string id = Controls.Add(Batch, first, "rich-text", "CustomerName").Controls[0].Id;
        Controls.Add(Batch, second, "rich-text", "CustomerName");

        var filled = Controls.SetText(Batch, "Only here", null, id);

        Assert.Equal(1, filled.AffectedCount);
        Assert.Equal("Only here", filled.Controls[0].Text);
        Assert.Contains(Controls.Get(Batch, "CustomerName").Controls, c => c.Text != "Only here");
    }

    [Fact]
    public void ContentControl_DeleteWithContentsRemovesTheText()
    {
        int index = Paragraphs.Add(Batch, "Draft note to remove.").Paragraph!.Index;
        Controls.Add(Batch, index, "rich-text", "Draft");

        Controls.Delete(Batch, "Draft", null, deleteContents: true);

        Assert.DoesNotContain("Draft note to remove.", Texts.Get(Batch).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentControl_ListCanBeFilteredByTag()
    {
        int first = Paragraphs.Add(Batch, "Customer.").Paragraph!.Index;
        int second = Paragraphs.Add(Batch, "Project.").Paragraph!.Index;

        Controls.Add(Batch, first, "rich-text", "CustomerName");
        Controls.Add(Batch, second, "rich-text", "ProjectName");

        Assert.Equal(2, Controls.List(Batch).TotalCount);
        Assert.Equal(1, Controls.List(Batch, "ProjectName").TotalCount);
    }

    [Fact]
    public void ContentControl_UnknownTagIsRejected()
    {
        Paragraphs.Add(Batch, "Nothing tagged here.");

        Assert.Throws<ArgumentException>(() => Controls.Get(Batch, "Missing"));
        Assert.Throws<ArgumentException>(() => Controls.SetText(Batch, "x", "Missing"));
        Assert.Throws<ArgumentException>(() => Controls.Delete(Batch, "Missing"));
    }

    [Fact]
    public void ContentControl_ParagraphBeyondTheDocumentIsRejected()
    {
        Paragraphs.Add(Batch, "One paragraph.");

        Assert.Throws<ArgumentOutOfRangeException>(() => Controls.Add(Batch, 999));
    }

    [Fact]
    public void ContentControl_PlainTextControlAcceptsText()
    {
        int index = Paragraphs.Add(Batch, "Plain value.").Paragraph!.Index;

        Controls.Add(Batch, index, "text", "Reference");
        var filled = Controls.SetText(Batch, "INV-2025-0042", "Reference");

        Assert.Equal("text", filled.Controls[0].Type);
        Assert.Equal("INV-2025-0042", filled.Controls[0].Text);
    }

    [Fact]
    public void ContentControl_ListReportsAnEmptyDocumentWithoutFailing()
    {
        var list = Controls.List(Batch);

        Assert.True(list.Success);
        Assert.Equal(0, list.TotalCount);
        Assert.Empty(list.Controls);
    }

    public void Dispose()
    {
        _sessions.Dispose();

        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
            // A locked temp file must not fail the test run.
        }
    }
}

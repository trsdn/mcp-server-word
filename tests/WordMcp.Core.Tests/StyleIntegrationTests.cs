using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Style;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the style commands. They use their own document because creating and
/// deleting styles changes the document-wide style collection.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class StyleIntegrationTests : IDisposable
{
    private static readonly StyleCommands Styles = new();
    private static readonly ParagraphCommands Paragraphs = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public StyleIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-style-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "styles.docx")).SessionId;
    }

    [Fact]
    public void Style_ListReportsFewerStylesWhenFilteringByUse()
    {
        var used = Styles.List(Batch);
        var all = Styles.List(Batch, inUseOnly: false);

        Assert.Equal(all.TotalCount, used.TotalCount);
        Assert.True(all.ReturnedCount > used.ReturnedCount);
        Assert.Equal(all.ReturnedCount, all.Styles.Count);
        Assert.All(used.Styles, style => Assert.True(style.InUse));
    }

    [Fact]
    public void Style_ListReportsEnglishNamesForBuiltInStyles()
    {
        var all = Styles.List(Batch, inUseOnly: false);

        var heading = Assert.Single(all.Styles, s => s.EnglishName == "Heading 1");

        Assert.True(heading.BuiltIn);
        Assert.Equal("paragraph", heading.Type);
        Assert.False(string.IsNullOrWhiteSpace(heading.Name));
    }

    [Fact]
    public void Style_ListFiltersByStyleType()
    {
        var characters = Styles.List(Batch, inUseOnly: false, styleType: "character");

        Assert.NotEmpty(characters.Styles);
        Assert.All(characters.Styles, style => Assert.Equal("character", style.Type));
    }

    [Fact]
    public void Style_CreateModifyApplyDeleteRoundTrip()
    {
        var created = Styles.Create(Batch, "Callout", "paragraph", "Normal");

        Assert.NotNull(created.Style);
        Assert.Equal("Callout", created.Style!.Name);
        Assert.False(created.Style.BuiltIn);
        Assert.Equal("paragraph", created.Style.Type);

        var modified = Styles.Modify(
            Batch,
            "Callout",
            fontName: "Calibri",
            fontSize: 14,
            bold: true,
            color: "#C00000",
            alignment: "center",
            spaceAfter: 12);

        Assert.Equal("Calibri", modified.Style!.FontName);
        Assert.Equal(14, modified.Style.FontSize);

        // Applying the style is the real proof that Word accepted it as a usable paragraph style.
        Paragraphs.Add(Batch, "Important note", style: "Callout");

        var inUse = Styles.List(Batch);
        Assert.Contains(inUse.Styles, style => style.Name == "Callout");

        // The style cannot be deleted while a paragraph still uses it in a way that matters, so the
        // paragraph is reset first.
        Paragraphs.SetStyle(Batch, 1, "Normal");

        var deleted = Styles.Delete(Batch, "Callout");
        Assert.Contains("deleted", deleted.Message, StringComparison.OrdinalIgnoreCase);

        Assert.DoesNotContain(Styles.List(Batch, inUseOnly: false).Styles, style => style.Name == "Callout");
    }

    [Fact]
    public void Style_CreateRejectsExistingName()
    {
        Styles.Create(Batch, "Duplicate");

        Assert.Throws<ArgumentException>(() => Styles.Create(Batch, "Duplicate"));
    }

    [Fact]
    public void Style_ModifyRejectsUnknownStyle()
        => Assert.Throws<ArgumentException>(() => Styles.Modify(Batch, "No Such Style", bold: true));

    [Fact]
    public void Style_ModifyAcceptsEnglishNameOfBuiltInStyle()
    {
        var modified = Styles.Modify(Batch, "Heading 1", fontSize: 20);

        Assert.Equal(20, modified.Style!.FontSize);
        Assert.True(modified.Style.BuiltIn);
        Assert.Equal("Heading 1", modified.Style.EnglishName);
    }

    [Fact]
    public void Style_DeleteRefusesBuiltInStyle()
        => Assert.Throws<InvalidOperationException>(() => Styles.Delete(Batch, "Heading 1"));

    [Fact]
    public void Style_DeleteRejectsUnknownStyle()
        => Assert.Throws<ArgumentException>(() => Styles.Delete(Batch, "No Such Style"));

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

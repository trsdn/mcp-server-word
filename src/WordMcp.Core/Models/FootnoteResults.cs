namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a single footnote or endnote.
/// </summary>
public sealed class FootnoteInfo
{
    /// <summary>Gets or sets the 1-based index within its collection.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the note kind, either <c>footnote</c> or <c>endnote</c>.</summary>
    public string Kind { get; set; } = "footnote";

    /// <summary>Gets or sets the reference mark as Word renders it, for example <c>1</c> or <c>*</c>.</summary>
    public string Reference { get; set; } = string.Empty;

    /// <summary>Gets or sets the 1-based index of the body paragraph the mark sits in.</summary>
    public int ParagraphIndex { get; set; }

    /// <summary>Gets or sets the note text, shortened for the listing.</summary>
    public string Text { get; set; } = string.Empty;
}

/// <summary>
/// The footnotes or endnotes of a document.
/// </summary>
public sealed class FootnoteListResult : ResultBase
{
    /// <summary>Gets or sets the note kind that was listed.</summary>
    public string Kind { get; set; } = "footnote";

    /// <summary>Gets or sets the number of notes of that kind.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the notes.</summary>
    public IReadOnlyList<FootnoteInfo> Notes { get; set; } = [];
}

/// <summary>
/// Result of an operation on a single footnote or endnote.
/// </summary>
public sealed class FootnoteResult : ResultBase
{
    /// <summary>Gets or sets the note the operation acted on.</summary>
    public FootnoteInfo? Note { get; set; }

    /// <summary>Gets or sets the number of notes of that kind left in the document.</summary>
    public int TotalCount { get; set; }
}

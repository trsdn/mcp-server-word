namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a single paragraph.
/// </summary>
public sealed class ParagraphInfo
{
    /// <summary>Gets or sets the 1-based paragraph index.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the paragraph text without the trailing paragraph mark.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Gets or sets the applied style name.</summary>
    public string Style { get; set; } = string.Empty;

    /// <summary>Gets or sets the paragraph alignment (left, center, right, justify).</summary>
    public string Alignment { get; set; } = string.Empty;

    /// <summary>Gets or sets the outline level, where 1-9 are heading levels and 10 is body text.</summary>
    public int OutlineLevel { get; set; }
}

/// <summary>
/// A page of paragraphs.
/// </summary>
public sealed class ParagraphListResult : ResultBase
{
    /// <summary>Gets or sets the total number of paragraphs in the document.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the returned paragraphs.</summary>
    public IReadOnlyList<ParagraphInfo> Paragraphs { get; set; } = [];
}

/// <summary>
/// Result of an operation that targets a single paragraph.
/// </summary>
public sealed class ParagraphResult : ResultBase
{
    /// <summary>Gets or sets the affected paragraph.</summary>
    public ParagraphInfo? Paragraph { get; set; }

    /// <summary>Gets or sets the total number of paragraphs after the operation.</summary>
    public int TotalCount { get; set; }
}

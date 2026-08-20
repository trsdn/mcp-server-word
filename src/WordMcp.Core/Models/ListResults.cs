namespace WordMcp.Core.Models;

/// <summary>
/// List formatting of a single paragraph.
/// </summary>
public sealed class ListParagraphInfo
{
    /// <summary>Gets or sets the 1-based paragraph index.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the paragraph text.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the kind of list: <c>bullet</c>, <c>number</c>, <c>outline-number</c> or
    /// <c>none</c> when the paragraph carries no list formatting.
    /// </summary>
    public string ListType { get; set; } = "none";

    /// <summary>Gets or sets the 1-based list level, or 0 when the paragraph is not in a list.</summary>
    public int Level { get; set; }

    /// <summary>
    /// Gets or sets the bullet or number Word renders in front of the paragraph, such as
    /// <c>2.</c> or <c>a)</c>. Empty when the paragraph is not in a list.
    /// </summary>
    public string ListLabel { get; set; } = string.Empty;
}

/// <summary>
/// Result of an operation that changes list formatting.
/// </summary>
public sealed class ListResult : ResultBase
{
    /// <summary>Gets or sets the paragraphs the operation touched.</summary>
    public IReadOnlyList<ListParagraphInfo> Paragraphs { get; set; } = [];

    /// <summary>Gets or sets the number of paragraphs the operation touched.</summary>
    public int UpdatedCount { get; set; }

    /// <summary>Gets or sets the number of paragraphs in the document.</summary>
    public int TotalCount { get; set; }
}

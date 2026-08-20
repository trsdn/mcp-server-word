namespace WordMcp.Core.Models;

/// <summary>
/// Content of a single header or footer.
/// </summary>
public sealed class HeaderFooterInfo
{
    /// <summary>Gets or sets the 1-based index of the owning section.</summary>
    public int SectionIndex { get; set; }

    /// <summary>Gets or sets the kind, either <c>header</c> or <c>footer</c>.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the type: <c>primary</c>, <c>first-page</c> or <c>even-pages</c>.
    /// </summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>Gets or sets the text content.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets whether this header or footer is inherited from the previous section.
    /// </summary>
    /// <remarks>
    /// A linked header shows the text of the previous section and cannot be changed on its own;
    /// writing to it unlinks it first.
    /// </remarks>
    public bool LinkedToPrevious { get; set; }

    /// <summary>
    /// Gets or sets whether the type is active. A <c>first-page</c> header is only rendered when
    /// the section has <c>DifferentFirstPage</c> switched on.
    /// </summary>
    public bool IsActive { get; set; }
}

/// <summary>
/// Headers or footers matching a query.
/// </summary>
public sealed class HeaderFooterListResult : ResultBase
{
    /// <summary>Gets or sets the total number of entries returned.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the headers or footers.</summary>
    public IReadOnlyList<HeaderFooterInfo> HeadersFooters { get; set; } = [];
}

/// <summary>
/// Result of an operation that writes to headers or footers.
/// </summary>
public sealed class HeaderFooterResult : ResultBase
{
    /// <summary>Gets or sets the number of headers or footers that were changed.</summary>
    public int UpdatedCount { get; set; }

    /// <summary>Gets or sets the entries as they are after the operation.</summary>
    public IReadOnlyList<HeaderFooterInfo> HeadersFooters { get; set; } = [];
}

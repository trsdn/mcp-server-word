namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a single bookmark of a document.
/// </summary>
public sealed class BookmarkInfo
{
    /// <summary>Gets or sets the bookmark name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the 1-based index of the paragraph the bookmark starts in.</summary>
    public int ParagraphIndex { get; set; }

    /// <summary>Gets or sets the character offset the bookmark starts at.</summary>
    public int Start { get; set; }

    /// <summary>Gets or sets the character offset the bookmark ends at.</summary>
    public int End { get; set; }

    /// <summary>Gets or sets whether the bookmark marks a position rather than a span of text.</summary>
    public bool Empty { get; set; }

    /// <summary>Gets or sets the bookmarked text, shortened for the listing.</summary>
    public string Text { get; set; } = string.Empty;
}

/// <summary>
/// The bookmarks of a document.
/// </summary>
public sealed class BookmarkListResult : ResultBase
{
    /// <summary>Gets or sets the number of bookmarks in the document.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the bookmarks.</summary>
    public IReadOnlyList<BookmarkInfo> Bookmarks { get; set; } = [];
}

/// <summary>
/// Result of an operation on a single bookmark.
/// </summary>
public sealed class BookmarkResult : ResultBase
{
    /// <summary>Gets or sets the bookmark the operation acted on.</summary>
    public BookmarkInfo? Bookmark { get; set; }

    /// <summary>Gets or sets the number of bookmarks left in the document.</summary>
    public int TotalCount { get; set; }
}

/// <summary>
/// The full text of a bookmark.
/// </summary>
public sealed class BookmarkTextResult : ResultBase
{
    /// <summary>Gets or sets the bookmark name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the bookmarked text.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Gets or sets the length of the bookmarked text in characters.</summary>
    public int Length { get; set; }
}

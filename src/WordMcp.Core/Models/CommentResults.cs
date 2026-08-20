namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a single comment of a document.
/// </summary>
public sealed class CommentInfo
{
    /// <summary>Gets or sets the 1-based comment index.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the author name.</summary>
    public string Author { get; set; } = string.Empty;

    /// <summary>Gets or sets the author initials.</summary>
    public string Initial { get; set; } = string.Empty;

    /// <summary>Gets or sets the date the comment was written.</summary>
    public DateTime? Date { get; set; }

    /// <summary>Gets or sets the comment text.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Gets or sets the document text the comment refers to.</summary>
    public string ScopeText { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets whether the comment is marked as resolved, or null on a Word version that does
    /// not support resolving comments.
    /// </summary>
    public bool? Resolved { get; set; }

    /// <summary>
    /// Gets or sets the 1-based index of the paragraph the comment refers to, or 0 when it could
    /// not be determined.
    /// </summary>
    public int ParagraphIndex { get; set; }
}

/// <summary>
/// The comments of a document.
/// </summary>
public sealed class CommentListResult : ResultBase
{
    /// <summary>Gets or sets the number of comments in the document.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the comments.</summary>
    public IReadOnlyList<CommentInfo> Comments { get; set; } = [];
}

/// <summary>
/// Result of an operation that adds, resolves or removes a comment.
/// </summary>
public sealed class CommentResult : ResultBase
{
    /// <summary>Gets or sets the affected comment, when it still exists.</summary>
    public CommentInfo? Comment { get; set; }

    /// <summary>Gets or sets the number of comments left in the document.</summary>
    public int TotalCount { get; set; }
}

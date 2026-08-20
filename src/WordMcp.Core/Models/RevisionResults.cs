namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a single tracked change of a document.
/// </summary>
public sealed class RevisionInfo
{
    /// <summary>Gets or sets the 1-based revision index.</summary>
    public int Index { get; set; }

    /// <summary>
    /// Gets or sets the kind of change: <c>insert</c>, <c>delete</c>, <c>format</c>,
    /// <c>moved-from</c>, <c>moved-to</c> or <c>other</c>.
    /// </summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>Gets or sets the author of the change.</summary>
    public string Author { get; set; } = string.Empty;

    /// <summary>Gets or sets the date of the change.</summary>
    public DateTime? Date { get; set; }

    /// <summary>Gets or sets the affected text.</summary>
    public string Text { get; set; } = string.Empty;
}

/// <summary>
/// The tracked changes of a document.
/// </summary>
public sealed class RevisionListResult : ResultBase
{
    /// <summary>Gets or sets the number of tracked changes in the document.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets whether change tracking is currently on.</summary>
    public bool TrackingEnabled { get; set; }

    /// <summary>Gets or sets the tracked changes.</summary>
    public IReadOnlyList<RevisionInfo> Revisions { get; set; } = [];
}

/// <summary>
/// Result of an operation that accepts or rejects changes, or switches tracking.
/// </summary>
public sealed class RevisionResult : ResultBase
{
    /// <summary>Gets or sets the number of changes the operation handled.</summary>
    public int HandledCount { get; set; }

    /// <summary>Gets or sets the number of tracked changes left in the document.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets whether change tracking is on after the operation.</summary>
    public bool TrackingEnabled { get; set; }
}

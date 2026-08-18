namespace WordMcp.Core.Models;

/// <summary>
/// Text content of a document or a range within it.
/// </summary>
public sealed class TextResult : ResultBase
{
    /// <summary>Gets or sets the text content.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Gets or sets the start character position of the returned range.</summary>
    public int Start { get; set; }

    /// <summary>Gets or sets the end character position of the returned range.</summary>
    public int End { get; set; }

    /// <summary>Gets or sets a value indicating whether the text was truncated by a length limit.</summary>
    public bool Truncated { get; set; }
}

/// <summary>
/// A single match of a text search.
/// </summary>
public sealed class TextMatch
{
    /// <summary>Gets or sets the start character position of the match.</summary>
    public int Start { get; set; }

    /// <summary>Gets or sets the end character position of the match.</summary>
    public int End { get; set; }

    /// <summary>Gets or sets the surrounding text of the match.</summary>
    public string Context { get; set; } = string.Empty;
}

/// <summary>
/// Result of a find operation.
/// </summary>
public sealed class FindResult : ResultBase
{
    /// <summary>Gets or sets the search term.</summary>
    public string SearchText { get; set; } = string.Empty;

    /// <summary>Gets or sets the total number of matches found.</summary>
    public int MatchCount { get; set; }

    /// <summary>Gets or sets the matches, limited by the requested maximum.</summary>
    public IReadOnlyList<TextMatch> Matches { get; set; } = [];
}

/// <summary>
/// Result of a replace operation.
/// </summary>
public sealed class ReplaceResult : ResultBase
{
    /// <summary>Gets or sets the search term.</summary>
    public string SearchText { get; set; } = string.Empty;

    /// <summary>Gets or sets the replacement text.</summary>
    public string ReplaceText { get; set; } = string.Empty;

    /// <summary>Gets or sets the number of replacements performed.</summary>
    public int ReplacementCount { get; set; }
}

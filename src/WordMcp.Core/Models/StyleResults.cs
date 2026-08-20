namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a single style of a document.
/// </summary>
public sealed class StyleInfo
{
    /// <summary>
    /// Gets or sets the name Word reports, which is localized on a non-English installation.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the English name of a built-in style, which is what clients should pass back.
    /// Empty for custom styles, whose name is already language independent.
    /// </summary>
    public string? EnglishName { get; set; }

    /// <summary>
    /// Gets or sets the kind of style: <c>paragraph</c>, <c>character</c>, <c>table</c> or <c>list</c>.
    /// </summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>Gets or sets whether Word ships the style.</summary>
    public bool BuiltIn { get; set; }

    /// <summary>Gets or sets whether the style is applied or modified in this document.</summary>
    public bool InUse { get; set; }

    /// <summary>Gets or sets the style this one inherits from, when there is one.</summary>
    public string? BaseStyle { get; set; }

    /// <summary>Gets or sets the font name, for paragraph and character styles.</summary>
    public string? FontName { get; set; }

    /// <summary>Gets or sets the font size in points, for paragraph and character styles.</summary>
    public double? FontSize { get; set; }
}

/// <summary>
/// The styles of a document.
/// </summary>
public sealed class StyleListResult : ResultBase
{
    /// <summary>Gets or sets the number of styles the document defines.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the number of styles returned after filtering.</summary>
    public int ReturnedCount { get; set; }

    /// <summary>Gets or sets the styles.</summary>
    public IReadOnlyList<StyleInfo> Styles { get; set; } = [];
}

/// <summary>
/// Result of an operation that creates, changes or removes a style.
/// </summary>
public sealed class StyleResult : ResultBase
{
    /// <summary>Gets or sets the affected style, when it still exists.</summary>
    public StyleInfo? Style { get; set; }
}

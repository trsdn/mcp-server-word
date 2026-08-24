namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for a single content control.
/// </summary>
public sealed class ContentControlInfo
{
    /// <summary>Gets or sets the 1-based index in the document's content control collection.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the Word-assigned identifier, stable for the lifetime of the control.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Gets or sets the tag, the machine-readable name a template uses.</summary>
    public string Tag { get; set; } = string.Empty;

    /// <summary>Gets or sets the title shown to the user in Word.</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>Gets or sets the control type, for example <c>rich-text</c> or <c>checkbox</c>.</summary>
    public string Type { get; set; } = string.Empty;

    /// <summary>Gets or sets the current text, shortened for listings.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Gets or sets a value indicating whether the content is locked against editing.</summary>
    public bool LockContents { get; set; }

    /// <summary>Gets or sets a value indicating whether the control itself cannot be deleted.</summary>
    public bool LockContentControl { get; set; }

    /// <summary>Gets or sets a value indicating whether a checkbox control is ticked.</summary>
    public bool? Checked { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the control still shows its placeholder rather than
    /// entered content.
    /// </summary>
    public bool ShowingPlaceholderText { get; set; }
}

/// <summary>
/// The content controls of a document.
/// </summary>
public sealed class ContentControlListResult : ResultBase
{
    /// <summary>Gets or sets the number of content controls found.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the content controls.</summary>
    public IReadOnlyList<ContentControlInfo> Controls { get; set; } = [];
}

/// <summary>
/// Result of an operation on one or more content controls.
/// </summary>
public sealed class ContentControlResult : ResultBase
{
    /// <summary>Gets or sets the controls the operation acted on.</summary>
    public IReadOnlyList<ContentControlInfo> Controls { get; set; } = [];

    /// <summary>Gets or sets the number of controls the operation acted on.</summary>
    public int AffectedCount { get; set; }
}

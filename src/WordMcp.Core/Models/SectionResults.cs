namespace WordMcp.Core.Models;

/// <summary>
/// Page setup of a section. All measurements are in points (72 pt = 1 inch), which is the unit
/// Word itself uses.
/// </summary>
public sealed class PageSetupInfo
{
    /// <summary>Gets or sets the top margin in points.</summary>
    public double TopMargin { get; set; }

    /// <summary>Gets or sets the bottom margin in points.</summary>
    public double BottomMargin { get; set; }

    /// <summary>Gets or sets the left margin in points.</summary>
    public double LeftMargin { get; set; }

    /// <summary>Gets or sets the right margin in points.</summary>
    public double RightMargin { get; set; }

    /// <summary>Gets or sets the page width in points.</summary>
    public double PageWidth { get; set; }

    /// <summary>Gets or sets the page height in points.</summary>
    public double PageHeight { get; set; }

    /// <summary>Gets or sets the orientation, either <c>portrait</c> or <c>landscape</c>.</summary>
    public string Orientation { get; set; } = string.Empty;
}

/// <summary>
/// Metadata for a section of the document.
/// </summary>
public sealed class SectionInfo
{
    /// <summary>Gets or sets the 1-based section index.</summary>
    public int Index { get; set; }

    /// <summary>
    /// Gets or sets how the section starts, for example <c>next-page</c> or <c>continuous</c>.
    /// </summary>
    public string StartType { get; set; } = string.Empty;

    /// <summary>Gets or sets the page setup of the section.</summary>
    public PageSetupInfo PageSetup { get; set; } = new();

    /// <summary>Gets or sets whether the first page has its own header and footer.</summary>
    public bool DifferentFirstPage { get; set; }

    /// <summary>Gets or sets whether odd and even pages have separate headers and footers.</summary>
    public bool DifferentOddEvenPages { get; set; }
}

/// <summary>
/// All sections of a document.
/// </summary>
public sealed class SectionListResult : ResultBase
{
    /// <summary>Gets or sets the total number of sections.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the sections.</summary>
    public IReadOnlyList<SectionInfo> Sections { get; set; } = [];
}

/// <summary>
/// Result of an operation that adds or changes a section.
/// </summary>
public sealed class SectionResult : ResultBase
{
    /// <summary>Gets or sets the affected section, when the operation targeted a single one.</summary>
    public SectionInfo? Section { get; set; }

    /// <summary>Gets or sets the total number of sections after the operation.</summary>
    public int TotalCount { get; set; }
}

namespace WordMcp.Core.Models;

/// <summary>
/// A rendered page image.
/// </summary>
public sealed class ScreenshotResult : ResultBase
{
    /// <summary>Gets or sets the 1-based page number that was rendered.</summary>
    public int Page { get; set; }

    /// <summary>Gets or sets the number of pages the document has.</summary>
    public int PageCount { get; set; }

    /// <summary>Gets or sets the path of the written PNG file.</summary>
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the image width in pixels.</summary>
    public int Width { get; set; }

    /// <summary>Gets or sets the image height in pixels.</summary>
    public int Height { get; set; }

    /// <summary>Gets or sets the resolution the page was rendered at.</summary>
    public int Dpi { get; set; }

    /// <summary>Gets or sets the size of the PNG file in bytes.</summary>
    public long FileSizeBytes { get; set; }

    /// <summary>
    /// Gets or sets the PNG as a base64 string. Only filled when the caller asked for it, because
    /// an inline image costs far more context than the path does.
    /// </summary>
    public string? ImageBase64 { get; set; }
}

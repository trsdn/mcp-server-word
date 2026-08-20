namespace WordMcp.Core.Models;

/// <summary>
/// Metadata for an inline image in the document.
/// </summary>
public sealed class ImageInfo
{
    /// <summary>Gets or sets the 1-based image index.</summary>
    public int Index { get; set; }

    /// <summary>Gets or sets the width in points.</summary>
    public double Width { get; set; }

    /// <summary>Gets or sets the height in points.</summary>
    public double Height { get; set; }

    /// <summary>Gets or sets the alternative text used by screen readers.</summary>
    public string AltText { get; set; } = string.Empty;

    /// <summary>Gets or sets whether resizing keeps the aspect ratio.</summary>
    public bool LockAspectRatio { get; set; }

    /// <summary>Gets or sets whether the image is linked to a file instead of embedded.</summary>
    public bool IsLinked { get; set; }
}

/// <summary>
/// All inline images of a document.
/// </summary>
public sealed class ImageListResult : ResultBase
{
    /// <summary>Gets or sets the total number of inline images.</summary>
    public int TotalCount { get; set; }

    /// <summary>Gets or sets the images.</summary>
    public IReadOnlyList<ImageInfo> Images { get; set; } = [];
}

/// <summary>
/// Result of an operation that modifies an image.
/// </summary>
public sealed class ImageResult : ResultBase
{
    /// <summary>Gets or sets the affected image.</summary>
    public ImageInfo? Image { get; set; }

    /// <summary>Gets or sets the total number of inline images after the operation.</summary>
    public int TotalCount { get; set; }
}

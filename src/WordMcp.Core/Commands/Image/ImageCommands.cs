using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Image;

/// <summary>
/// Word COM implementation of <see cref="IImageCommands"/>.
/// </summary>
public sealed class ImageCommands : IImageCommands
{
    /// <inheritdoc />
    public ImageListResult List(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic shapes = ctx.Document.InlineShapes;
            int total = (int)shapes.Count;

            var list = new List<ImageInfo>(total);
            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();
                list.Add(Describe(shapes[i], i));
            }

            return new ImageListResult
            {
                TotalCount = total,
                Images = list
            };
        });
    }

    /// <inheritdoc />
    public ImageResult Insert(
        IWordBatch batch,
        string imagePath,
        int? paragraphIndex = null,
        double? width = null,
        double? height = null,
        string? caption = null,
        string? altText = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);

        string fullPath = ValidateImagePath(imagePath);

        if (paragraphIndex is int p)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(p, 1);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic range = ResolveInsertRange(doc, paragraphIndex);

            dynamic shape = doc.InlineShapes.AddPicture(fullPath, false, true, range);

            ApplySize(shape, width, height, scalePercent: null, lockAspectRatio: true);

            if (!string.IsNullOrEmpty(altText))
            {
                shape.AlternativeText = altText;
            }

            if (!string.IsNullOrWhiteSpace(caption))
            {
                InsertCaption(shape, caption);
            }

            int index = IndexOf(doc, shape);

            return new ImageResult
            {
                Image = Describe(shape, index),
                TotalCount = (int)doc.InlineShapes.Count,
                Message = $"Image inserted at index {index}."
            };
        });
    }

    /// <inheritdoc />
    public ImageResult Resize(
        IWordBatch batch,
        int index,
        double? width = null,
        double? height = null,
        double? scalePercent = null,
        bool lockAspectRatio = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        if (width is null && height is null && scalePercent is null)
        {
            throw new ArgumentException(
                "Specify width, height or scalePercent.", nameof(scalePercent));
        }

        if (scalePercent is double s && s <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(scalePercent), scalePercent, "scalePercent must be greater than 0.");
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic shape = GetShape(doc, index);

            ApplySize(shape, width, height, scalePercent, lockAspectRatio);

            return new ImageResult
            {
                Image = Describe(shape, index),
                TotalCount = (int)doc.InlineShapes.Count,
                Message = $"Image {index} resized."
            };
        });
    }

    /// <inheritdoc />
    public ImageResult Replace(IWordBatch batch, int index, string imagePath, bool keepSize = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        string fullPath = ValidateImagePath(imagePath);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic old = GetShape(doc, index);

            double oldWidth = (double)(float)old.Width;
            double oldHeight = (double)(float)old.Height;
            string oldAltText = ReadAltText(old);

            // Anchor the replacement on the old picture's range so it lands in the same spot.
            dynamic range = old.Range;
            old.Delete();

            dynamic shape = doc.InlineShapes.AddPicture(fullPath, false, true, range);

            if (keepSize)
            {
                shape.LockAspectRatio = ComInteropConstants.MsoFalse;
                shape.Width = (float)oldWidth;
                shape.Height = (float)oldHeight;
            }

            if (!string.IsNullOrEmpty(oldAltText))
            {
                shape.AlternativeText = oldAltText;
            }

            int newIndex = IndexOf(doc, shape);

            return new ImageResult
            {
                Image = Describe(shape, newIndex),
                TotalCount = (int)doc.InlineShapes.Count,
                Message = $"Image {newIndex} replaced with '{Path.GetFileName(fullPath)}'."
            };
        });
    }

    /// <inheritdoc />
    public ImageResult Delete(IWordBatch batch, int index)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            GetShape(doc, index).Delete();

            return new ImageResult
            {
                TotalCount = (int)doc.InlineShapes.Count,
                Message = $"Image {index} deleted."
            };
        });
    }

    /// <inheritdoc />
    public ImageResult SetAltText(IWordBatch batch, int index, string altText)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(altText);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic shape = GetShape(doc, index);
            shape.AlternativeText = altText;

            return new ImageResult
            {
                Image = Describe(shape, index),
                TotalCount = (int)doc.InlineShapes.Count,
                Message = $"Alternative text of image {index} updated."
            };
        });
    }

    /// <summary>
    /// Validates that the path points at an existing image file Word can insert.
    /// </summary>
    private static string ValidateImagePath(string imagePath)
    {
        string fullPath = Path.GetFullPath(imagePath);

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Image file not found: '{fullPath}'.", fullPath);
        }

        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (!ComInteropConstants.SupportedImageExtensions.Contains(extension))
        {
            throw new ArgumentException(
                $"Unsupported image format '{extension}'. Supported: " +
                string.Join(", ", ComInteropConstants.SupportedImageExtensions) + ".",
                nameof(imagePath));
        }

        return fullPath;
    }

    /// <summary>
    /// Returns the range an inserted image should be anchored on.
    /// </summary>
    private static dynamic ResolveInsertRange(dynamic doc, int? paragraphIndex)
    {
        if (paragraphIndex is not int index)
        {
            dynamic end = doc.Content;
            end.Collapse(ComInteropConstants.WdCollapseEnd);
            end.InsertParagraphAfter();

            dynamic target = doc.Content;
            target.Collapse(ComInteropConstants.WdCollapseEnd);
            return target;
        }

        int total = (int)doc.Paragraphs.Count;
        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(paragraphIndex),
                $"Paragraph {index} does not exist. The document has {total} paragraph(s).");
        }

        dynamic range = doc.Paragraphs[index].Range;
        range.Collapse(ComInteropConstants.WdCollapseStart);
        return range;
    }

    /// <summary>
    /// Adds a caption below the image. The range of an inline shape covers the shape character
    /// itself and rejects edits, so this uses Word's own caption machinery (which also numbers the
    /// caption and makes it available to a table of figures) and only falls back to a plain
    /// paragraph if that is unavailable.
    /// </summary>
    private static void InsertCaption(dynamic shape, string caption)
    {
        dynamic range = shape.Range;

        try
        {
            range.InsertCaption(
                ComInteropConstants.WdCaptionFigure,
                " " + caption,
                Type.Missing,
                ComInteropConstants.WdCaptionPositionBelow,
                false);
            return;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Fall through to the manual paragraph below.
        }

        dynamic paragraph = range.Paragraphs[1];
        dynamic anchor = paragraph.Range;
        anchor.InsertParagraphAfter();

        dynamic captionRange = anchor.Paragraphs[anchor.Paragraphs.Count].Range;
        captionRange.Text = caption;

        try
        {
            captionRange.Style = WordStyles.Resolve("Caption");
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Without the Caption style the text is still there, just unstyled.
        }
    }

    private static void ApplySize(
        dynamic shape,
        double? width,
        double? height,
        double? scalePercent,
        bool lockAspectRatio)
    {
        if (scalePercent is double percent)
        {
            // ScaleWidth/ScaleHeight are relative to the original size, so scaling repeatedly
            // would not compound. Deriving from the current size keeps it predictable.
            double currentWidth = (double)(float)shape.Width;
            double currentHeight = (double)(float)shape.Height;

            shape.LockAspectRatio = ComInteropConstants.MsoFalse;
            shape.Width = (float)(currentWidth * percent / 100d);
            shape.Height = (float)(currentHeight * percent / 100d);
            return;
        }

        if (width is null && height is null)
        {
            return;
        }

        shape.LockAspectRatio = lockAspectRatio
            ? ComInteropConstants.MsoTrue
            : ComInteropConstants.MsoFalse;

        // With a locked aspect ratio Word derives the second dimension, so setting width last
        // would silently override an explicit height.
        if (width is double w)
        {
            shape.Width = (float)w;
        }

        if (height is double h)
        {
            shape.Height = (float)h;
        }
    }

    private static dynamic GetShape(dynamic doc, int index)
    {
        int total = (int)doc.InlineShapes.Count;
        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index), $"Image {index} does not exist. The document has {total} image(s).");
        }

        return doc.InlineShapes[index];
    }

    /// <summary>
    /// Finds the 1-based index of a shape by comparing range positions.
    /// </summary>
    /// <remarks>
    /// <c>InlineShapes</c> has no IndexOf, and an inserted picture is not necessarily the last
    /// element when it was anchored somewhere in the middle of the document.
    /// </remarks>
    private static int IndexOf(dynamic doc, dynamic shape)
    {
        int start = (int)shape.Range.Start;
        int total = (int)doc.InlineShapes.Count;

        for (int i = 1; i <= total; i++)
        {
            if ((int)doc.InlineShapes[i].Range.Start == start)
            {
                return i;
            }
        }

        return total;
    }

    private static string ReadAltText(dynamic shape)
    {
        try
        {
            return (string?)shape.AlternativeText ?? string.Empty;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            return string.Empty;
        }
    }

    private static ImageInfo Describe(dynamic shape, int index) => new()
    {
        Index = index,
        Width = Math.Round((double)(float)shape.Width, 2),
        Height = Math.Round((double)(float)shape.Height, 2),
        AltText = ReadAltText(shape),
        LockAspectRatio = (int)shape.LockAspectRatio == ComInteropConstants.MsoTrue,
        IsLinked = (int)shape.Type == ComInteropConstants.WdInlineShapeLinkedPicture
    };
}

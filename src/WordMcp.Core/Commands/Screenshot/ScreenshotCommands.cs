using PDFtoImage;
using SkiaSharp;
using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Screenshot;

/// <summary>
/// Word COM implementation of <see cref="IScreenshotCommands"/>.
/// </summary>
/// <remarks>
/// Word has no API that hands out a page as an image. The only route that preserves the real
/// layout is a fixed-format export of the single page followed by rasterizing that PDF, which is
/// what this class does.
/// </remarks>
public sealed class ScreenshotCommands : IScreenshotCommands
{
    private const int MinDpi = 36;
    private const int MaxDpi = 600;

    /// <inheritdoc />
    public ScreenshotResult Page(
        IWordBatch batch,
        int page = 1,
        string? outputPath = null,
        int dpi = 150,
        bool includeImage = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(page, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(dpi, MinDpi);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(dpi, MaxDpi);

        string target = ResolveOutputPath(outputPath, page);

        var (pdfPath, pageCount) = ExportPage(batch, page);

        try
        {
            var (width, height) = Rasterize(pdfPath, target, dpi);
            var info = new FileInfo(target);

            return new ScreenshotResult
            {
                Page = page,
                PageCount = pageCount,
                OutputPath = target,
                Width = width,
                Height = height,
                Dpi = dpi,
                FileSizeBytes = info.Exists ? info.Length : 0,
                ImageBase64 = includeImage ? Convert.ToBase64String(File.ReadAllBytes(target)) : null,
                Message = $"Page {page} of {pageCount} rendered at {dpi} dpi."
            };
        }
        finally
        {
            TryDelete(pdfPath);
        }
    }

    private static string ResolveOutputPath(string? outputPath, int page)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return Path.Combine(
                Path.GetTempPath(),
                $"wordmcp-page{page}-{Guid.NewGuid():N}.png");
        }

        string full = Path.GetFullPath(outputPath);
        string? directory = Path.GetDirectoryName(full);

        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return full;
    }

    /// <summary>
    /// Exports the single page as a PDF and reports how many pages the document has.
    /// </summary>
    private static (string PdfPath, int PageCount) ExportPage(IWordBatch batch, int page)
    {
        string pdfPath = Path.Combine(Path.GetTempPath(), $"wordmcp-page-{Guid.NewGuid():N}.pdf");

        int pageCount = batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;

            // Repaginate first: on a document that was only ever touched through automation the
            // page count is otherwise stale, and a page that Word thinks does not exist exports
            // as an empty PDF instead of failing.
            doc.Repaginate();
            int pages = (int)doc.ComputeStatistics(ComInteropConstants.WdStatisticPages);

            if (page > pages)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(page), $"Page {page} does not exist. The document has {pages} page(s).");
            }

            doc.ExportAsFixedFormat(
                pdfPath,
                ComInteropConstants.WdExportFormatPdf,
                false,
                ComInteropConstants.WdExportOptimizeForPrint,
                ComInteropConstants.WdExportFromTo,
                page,
                page,
                ComInteropConstants.WdExportDocumentContent);

            return pages;
        });

        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException(
                $"Word did not produce a PDF for page {page}. The document may be protected against exporting.");
        }

        return (pdfPath, pageCount);
    }

    private static (int Width, int Height) Rasterize(string pdfPath, string target, int dpi)
    {
        using var pdf = File.OpenRead(pdfPath);

        // The export contains exactly the one requested page, so the rasterizer always reads index 0.
        using SKBitmap bitmap = Conversion.ToImage(pdf, page: 0, options: new RenderOptions(Dpi: dpi));
        using SKData png = bitmap.Encode(SKEncodedImageFormat.Png, 100)
            ?? throw new InvalidOperationException("The rendered page could not be encoded as PNG.");

        using var output = File.Create(target);
        png.SaveTo(output);

        return (bitmap.Width, bitmap.Height);
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
            // A leftover temp file is not worth failing the call over.
        }
        catch (UnauthorizedAccessException)
        {
            // Same here.
        }
    }
}

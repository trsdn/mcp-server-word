using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Document;

/// <summary>
/// Word COM implementation of <see cref="IDocumentCommands"/>.
/// </summary>
public sealed class DocumentCommands : IDocumentCommands
{
    /// <inheritdoc />
    public DocumentInfoResult GetInfo(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;

            return new DocumentInfoResult
            {
                FilePath = ctx.DocumentPath,
                FileName = Path.GetFileName(ctx.DocumentPath),
                WordCount = (int)doc.ComputeStatistics(ComInteropConstants.WdStatisticWords),
                CharacterCount = (int)doc.ComputeStatistics(ComInteropConstants.WdStatisticCharacters),
                ParagraphCount = (int)doc.Paragraphs.Count,
                PageCount = (int)doc.ComputeStatistics(ComInteropConstants.WdStatisticPages),
                TableCount = (int)doc.Tables.Count,
                InlineShapeCount = (int)doc.InlineShapes.Count,
                SectionCount = (int)doc.Sections.Count,
                HasUnsavedChanges = !(bool)doc.Saved,
                IsReadOnly = (bool)doc.ReadOnly
            };
        });
    }

    /// <inheritdoc />
    public DocumentPropertiesResult GetProperties(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic properties = ctx.Document.BuiltInDocumentProperties;

            return new DocumentPropertiesResult
            {
                Title = ReadProperty(properties, "Title"),
                Author = ReadProperty(properties, "Author"),
                Subject = ReadProperty(properties, "Subject"),
                Keywords = ReadProperty(properties, "Keywords"),
                Comments = ReadProperty(properties, "Comments"),
                Company = ReadProperty(properties, "Company"),
                LastAuthor = ReadProperty(properties, "Last author")
            };
        });
    }

    /// <inheritdoc />
    public DocumentPropertiesResult SetProperties(
        IWordBatch batch,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        string? comments = null,
        string? company = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        batch.Execute((ctx, ct) =>
        {
            dynamic properties = ctx.Document.BuiltInDocumentProperties;

            WriteProperty(properties, "Title", title);
            WriteProperty(properties, "Author", author);
            WriteProperty(properties, "Subject", subject);
            WriteProperty(properties, "Keywords", keywords);
            WriteProperty(properties, "Comments", comments);
            WriteProperty(properties, "Company", company);
        });

        var result = GetProperties(batch);
        result.Message = "Document properties updated.";
        return result;
    }

    /// <inheritdoc />
    public ExportResult ExportPdf(IWordBatch batch, string outputPath)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var fullPath = Path.GetFullPath(outputPath);
        if (!string.Equals(Path.GetExtension(fullPath), ".pdf", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The output path for export-pdf must end in .pdf.", nameof(outputPath));
        }

        return SaveCopy(batch, fullPath, ComInteropConstants.WdFormatPdf);
    }

    /// <inheritdoc />
    public ExportResult SaveAs(IWordBatch batch, string outputPath)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var fullPath = Path.GetFullPath(outputPath);
        return SaveCopy(batch, fullPath, WordConversions.ToWdSaveFormat(fullPath));
    }

    private static ExportResult SaveCopy(IWordBatch batch, string fullPath, int wdFormat)
    {
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;

            if (wdFormat == ComInteropConstants.WdFormatPdf)
            {
                // wdExportFormatPDF = 17, OpenAfterExport = false. Leaves the open document untouched.
                doc.ExportAsFixedFormat(fullPath, 17, false);
            }
            else
            {
                // Word has no "save a copy in another format" API: SaveAs2 always re-points the open
                // document at the new file. Saving back to the original path restores the session so
                // later operations keep writing to the document the caller opened. Side effect: the
                // original document is persisted as part of save-as.
                string originalPath = ctx.DocumentPath;
                doc.SaveAs2(fullPath, wdFormat);
                doc.SaveAs2(originalPath, WordConversions.ToWdSaveFormat(originalPath));
            }
        });

        var info = new FileInfo(fullPath);
        return new ExportResult
        {
            OutputPath = fullPath,
            FileSizeBytes = info.Exists ? info.Length : 0,
            Message = $"Exported to {fullPath}"
        };
    }

    private static string? ReadProperty(dynamic properties, string name)
    {
        try
        {
            var value = properties[name].Value;
            return value?.ToString();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Property not present in this document.
            return null;
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            return null;
        }
    }

    private static void WriteProperty(dynamic properties, string name, string? value)
    {
        if (value == null)
            return;

        properties[name].Value = value;
    }
}

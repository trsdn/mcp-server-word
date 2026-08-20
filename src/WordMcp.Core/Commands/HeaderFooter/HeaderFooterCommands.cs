using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.HeaderFooter;

/// <summary>
/// Word COM implementation of <see cref="IHeaderFooterCommands"/>.
/// </summary>
public sealed class HeaderFooterCommands : IHeaderFooterCommands
{
    /// <inheritdoc />
    public HeaderFooterListResult Get(
        IWordBatch batch,
        string kind = "header",
        int? sectionIndex = null,
        string type = "primary")
    {
        ArgumentNullException.ThrowIfNull(batch);

        bool isHeader = ParseKind(kind);
        int wdType = WordConversions.ToWdHeaderFooterIndex(type);

        if (sectionIndex is int s)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(s, 1);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            var entries = new List<HeaderFooterInfo>();

            foreach (int index in ResolveSections(doc, sectionIndex))
            {
                ct.ThrowIfCancellationRequested();

                dynamic section = doc.Sections[index];
                entries.Add(Describe(section, index, isHeader, wdType));
            }

            return new HeaderFooterListResult
            {
                TotalCount = entries.Count,
                HeadersFooters = entries
            };
        });
    }

    /// <inheritdoc />
    public HeaderFooterResult Set(
        IWordBatch batch,
        string text,
        string kind = "header",
        int? sectionIndex = null,
        string type = "primary",
        string? alignment = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return Write(batch, text, kind, sectionIndex, type, alignment);
    }

    /// <inheritdoc />
    public HeaderFooterResult Clear(
        IWordBatch batch,
        string kind = "header",
        int? sectionIndex = null,
        string type = "primary")
    {
        ArgumentNullException.ThrowIfNull(batch);

        return Write(batch, string.Empty, kind, sectionIndex, type, alignment: null);
    }

    private static HeaderFooterResult Write(
        IWordBatch batch,
        string text,
        string kind,
        int? sectionIndex,
        string type,
        string? alignment)
    {
        bool isHeader = ParseKind(kind);
        int wdType = WordConversions.ToWdHeaderFooterIndex(type);
        int? wdAlignment = alignment is null ? null : WordConversions.ToWdAlignment(alignment);

        if (sectionIndex is int s)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(s, 1);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            var entries = new List<HeaderFooterInfo>();

            foreach (int index in ResolveSections(doc, sectionIndex))
            {
                ct.ThrowIfCancellationRequested();

                dynamic section = doc.Sections[index];

                EnableType(section, wdType);

                dynamic headerFooter = GetHeaderFooter(section, isHeader, wdType);

                // A linked header shows the previous section's text and cannot be edited on its
                // own, so writing to a specific section has to break the link first.
                if (index > 1 && (bool)headerFooter.LinkToPrevious)
                {
                    headerFooter.LinkToPrevious = false;
                }

                dynamic range = headerFooter.Range;
                range.Text = text;

                if (wdAlignment is int align)
                {
                    range.ParagraphFormat.Alignment = align;
                }

                entries.Add(Describe(section, index, isHeader, wdType));
            }

            string what = isHeader ? "Header" : "Footer";
            string action = text.Length == 0 ? "cleared" : "set";

            return new HeaderFooterResult
            {
                UpdatedCount = entries.Count,
                HeadersFooters = entries,
                Message = $"{what} ({type}) {action} for {entries.Count} section(s)."
            };
        });
    }

    /// <summary>
    /// Turns on the section switch that makes a first-page or even-pages header visible. Without
    /// it Word stores the text but never renders it, which looks like a silent failure.
    /// </summary>
    private static void EnableType(dynamic section, int wdType)
    {
        dynamic pageSetup = section.PageSetup;

        if (wdType == ComInteropConstants.WdHeaderFooterFirstPage)
        {
            pageSetup.DifferentFirstPageHeaderFooter = ComInteropConstants.MsoTrue;
        }
        else if (wdType == ComInteropConstants.WdHeaderFooterEvenPages)
        {
            pageSetup.OddAndEvenPagesHeaderFooter = ComInteropConstants.MsoTrue;
        }
    }

    private static IEnumerable<int> ResolveSections(dynamic doc, int? sectionIndex)
    {
        int total = (int)doc.Sections.Count;

        if (sectionIndex is not int index)
        {
            return Enumerable.Range(1, total);
        }

        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sectionIndex),
                $"Section {index} does not exist. The document has {total} section(s).");
        }

        return [index];
    }

    private static dynamic GetHeaderFooter(dynamic section, bool isHeader, int wdType)
        => isHeader ? section.Headers[wdType] : section.Footers[wdType];

    private static bool ParseKind(string kind)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(kind);

        return kind.Trim().ToLowerInvariant() switch
        {
            "header" or "top" => true,
            "footer" or "bottom" => false,
            _ => throw new ArgumentException(
                $"Unknown kind '{kind}'. Use header or footer.", nameof(kind))
        };
    }

    private static HeaderFooterInfo Describe(dynamic section, int index, bool isHeader, int wdType)
    {
        dynamic headerFooter = GetHeaderFooter(section, isHeader, wdType);
        dynamic pageSetup = section.PageSetup;

        bool isActive = wdType switch
        {
            ComInteropConstants.WdHeaderFooterFirstPage =>
                (int)pageSetup.DifferentFirstPageHeaderFooter != 0,
            ComInteropConstants.WdHeaderFooterEvenPages =>
                (int)pageSetup.OddAndEvenPagesHeaderFooter != 0,
            _ => true
        };

        return new HeaderFooterInfo
        {
            SectionIndex = index,
            Kind = isHeader ? "header" : "footer",
            Type = WordConversions.FromWdHeaderFooterIndex(wdType),
            Text = WordConversions.CleanRangeText((string?)headerFooter.Range.Text),
            LinkedToPrevious = index > 1 && (bool)headerFooter.LinkToPrevious,
            IsActive = isActive
        };
    }
}

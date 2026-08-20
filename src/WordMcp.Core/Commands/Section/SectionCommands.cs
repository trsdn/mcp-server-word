using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Section;

/// <summary>
/// Word COM implementation of <see cref="ISectionCommands"/>.
/// </summary>
public sealed class SectionCommands : ISectionCommands
{
    /// <inheritdoc />
    public SectionListResult List(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic sections = ctx.Document.Sections;
            int total = (int)sections.Count;

            var list = new List<SectionInfo>(total);
            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();
                list.Add(Describe(sections[i], i));
            }

            return new SectionListResult
            {
                TotalCount = total,
                Sections = list
            };
        });
    }

    /// <inheritdoc />
    public SectionResult Add(IWordBatch batch, string startType = "next-page", int? paragraphIndex = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        int breakType = WordConversions.ToWdSectionBreak(startType);

        if (paragraphIndex is int p)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(p, 1);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic range = ResolveBreakRange(doc, paragraphIndex);

            range.InsertBreak(breakType);

            int total = (int)doc.Sections.Count;

            // The break splits the section that contained the range, so the new section is the one
            // right after it rather than necessarily the last one.
            int newIndex = (int)range.Sections[1].Index;
            if (newIndex < total)
            {
                newIndex++;
            }

            return new SectionResult
            {
                Section = Describe(doc.Sections[newIndex], newIndex),
                TotalCount = total,
                Message = $"Section break inserted; the document now has {total} section(s)."
            };
        });
    }

    /// <inheritdoc />
    public SectionResult PageSetup(
        IWordBatch batch,
        int? sectionIndex = null,
        double? topMargin = null,
        double? bottomMargin = null,
        double? leftMargin = null,
        double? rightMargin = null,
        string? orientation = null,
        string? paperSize = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (sectionIndex is int s)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(s, 1);
        }

        ValidateMargin(topMargin, nameof(topMargin));
        ValidateMargin(bottomMargin, nameof(bottomMargin));
        ValidateMargin(leftMargin, nameof(leftMargin));
        ValidateMargin(rightMargin, nameof(rightMargin));

        if (topMargin is null && bottomMargin is null && leftMargin is null && rightMargin is null
            && orientation is null && paperSize is null)
        {
            throw new ArgumentException(
                "Specify at least one of topMargin, bottomMargin, leftMargin, rightMargin, "
                + "orientation or paperSize.",
                nameof(sectionIndex));
        }

        // Parsed up front so an unknown name fails before anything is written.
        int? wdOrientation = orientation is null ? null : WordConversions.ToWdOrientation(orientation);
        int? wdPaperSize = paperSize is null ? null : WordConversions.ToWdPaperSize(paperSize);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int total = (int)doc.Sections.Count;

            dynamic target = sectionIndex is int index
                ? GetSection(doc, index, total)
                : doc;

            dynamic pageSetup = target.PageSetup;

            // Paper size resets the margins, so it has to be applied before them.
            if (wdPaperSize is int paper)
            {
                pageSetup.PaperSize = paper;
            }

            if (wdOrientation is int orient)
            {
                pageSetup.Orientation = orient;
            }

            if (topMargin is double top)
            {
                pageSetup.TopMargin = (float)top;
            }

            if (bottomMargin is double bottom)
            {
                pageSetup.BottomMargin = (float)bottom;
            }

            if (leftMargin is double left)
            {
                pageSetup.LeftMargin = (float)left;
            }

            if (rightMargin is double right)
            {
                pageSetup.RightMargin = (float)right;
            }

            int described = sectionIndex ?? 1;

            return new SectionResult
            {
                Section = Describe(doc.Sections[described], described),
                TotalCount = total,
                Message = sectionIndex is null
                    ? $"Page setup applied to all {total} section(s)."
                    : $"Page setup applied to section {sectionIndex}."
            };
        });
    }

    private static void ValidateMargin(double? margin, string parameterName)
    {
        if (margin is double value && value < 0)
        {
            throw new ArgumentOutOfRangeException(parameterName, value, "Margins cannot be negative.");
        }
    }

    private static dynamic GetSection(dynamic doc, int index, int total)
    {
        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index), $"Section {index} does not exist. The document has {total} section(s).");
        }

        return doc.Sections[index];
    }

    /// <summary>
    /// Resolves the range the section break is inserted at. Without a paragraph the break goes to
    /// the end of the document.
    /// </summary>
    private static dynamic ResolveBreakRange(dynamic doc, int? paragraphIndex)
    {
        if (paragraphIndex is not int index)
        {
            dynamic content = doc.Content;
            content.Collapse(ComInteropConstants.WdCollapseEnd);
            return content;
        }

        int total = (int)doc.Paragraphs.Count;
        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(paragraphIndex),
                $"Paragraph {index} does not exist. The document has {total} paragraph(s).");
        }

        dynamic range = doc.Paragraphs[index].Range;
        range.Collapse(ComInteropConstants.WdCollapseEnd);
        return range;
    }

    private static SectionInfo Describe(dynamic section, int index)
    {
        dynamic pageSetup = section.PageSetup;

        return new SectionInfo
        {
            Index = index,
            StartType = WordConversions.FromWdSectionStart((int)pageSetup.SectionStart),
            DifferentFirstPage = (int)pageSetup.DifferentFirstPageHeaderFooter != 0,
            DifferentOddEvenPages = (int)pageSetup.OddAndEvenPagesHeaderFooter != 0,
            PageSetup = new PageSetupInfo
            {
                TopMargin = Round((double)pageSetup.TopMargin),
                BottomMargin = Round((double)pageSetup.BottomMargin),
                LeftMargin = Round((double)pageSetup.LeftMargin),
                RightMargin = Round((double)pageSetup.RightMargin),
                PageWidth = Round((double)pageSetup.PageWidth),
                PageHeight = Round((double)pageSetup.PageHeight),
                Orientation = WordConversions.FromWdOrientation((int)pageSetup.Orientation)
            }
        };
    }

    /// <summary>
    /// Word stores measurements as single-precision floats, so values come back as 56.70000076.
    /// </summary>
    private static double Round(double value) => Math.Round(value, 2);
}

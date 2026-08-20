using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Field;

/// <summary>
/// Word COM implementation of <see cref="IFieldCommands"/>.
/// </summary>
public sealed class FieldCommands : IFieldCommands
{
    /// <inheritdoc />
    public FieldListResult List(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic fields = ctx.Document.Fields;
            int total = (int)fields.Count;

            var list = new List<Models.FieldInfo>(total);
            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();
                list.Add(Describe(fields[i], i));
            }

            return new FieldListResult
            {
                TotalCount = total,
                Fields = list
            };
        });
    }

    /// <inheritdoc />
    public FieldResult InsertTableOfContents(
        IWordBatch batch,
        int? paragraphIndex = null,
        int upperHeadingLevel = 1,
        int lowerHeadingLevel = 3,
        bool includePageNumbers = true,
        bool useHyperlinks = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(upperHeadingLevel, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(lowerHeadingLevel, 9);

        if (lowerHeadingLevel < upperHeadingLevel)
        {
            throw new ArgumentException(
                $"lowerHeadingLevel ({lowerHeadingLevel}) must not be smaller than " +
                $"upperHeadingLevel ({upperHeadingLevel}).",
                nameof(lowerHeadingLevel));
        }

        if (paragraphIndex is int p)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(p, 1);
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic range = ResolveInsertRange(doc, paragraphIndex);

            dynamic toc = doc.TablesOfContents.Add(
                range,
                true,                 // UseHeadingStyles
                upperHeadingLevel,
                lowerHeadingLevel,
                false,                // UseFields
                Type.Missing,         // TableID
                true,                 // RightAlignPageNumbers
                includePageNumbers,
                Type.Missing,         // AddedStyles
                useHyperlinks);

            toc.Update();

            int entries = CountEntries(doc, toc);

            string message = entries == 0
                ? "Table of contents inserted but empty: no paragraph uses a heading style. " +
                  "Apply 'Heading 1' and friends via paragraph(add|set-style), then run field(update-toc)."
                : $"Table of contents inserted with {entries} entr{(entries == 1 ? "y" : "ies")}.";

            return new FieldResult
            {
                UpdatedCount = 1,
                EntryCount = entries,
                Message = message
            };
        });
    }

    /// <inheritdoc />
    public FieldResult UpdateTableOfContents(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic tocs = doc.TablesOfContents;
            int total = (int)tocs.Count;

            int entries = 0;
            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();

                dynamic toc = tocs[i];
                toc.Update();
                entries += CountEntries(doc, toc);
            }

            return new FieldResult
            {
                UpdatedCount = total,
                EntryCount = entries,
                Message = total == 0
                    ? "The document has no table of contents. Insert one with field(insert-toc)."
                    : $"{total} table(s) of contents updated, {entries} entr{(entries == 1 ? "y" : "ies")} in total."
            };
        });
    }

    /// <inheritdoc />
    public FieldResult UpdateAll(IWordBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;

            int updated = (int)doc.Fields.Count;
            doc.Fields.Update();

            // Page numbers live in headers and footers, which Document.Fields does not cover.
            foreach (dynamic section in doc.Sections)
            {
                ct.ThrowIfCancellationRequested();
                updated += UpdateHeaderFooterFields(section.Headers);
                updated += UpdateHeaderFooterFields(section.Footers);
            }

            return new FieldResult
            {
                UpdatedCount = updated,
                Message = $"{updated} field(s) updated."
            };
        });
    }

    /// <inheritdoc />
    public FieldResult InsertPageNumber(
        IWordBatch batch,
        string position = "footer",
        string alignment = "center",
        bool includeTotalPages = false)
    {
        ArgumentNullException.ThrowIfNull(batch);

        bool inHeader = ParsePosition(position);
        int wdAlignment = WordConversions.ToWdAlignment(alignment);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int inserted = 0;

            foreach (dynamic section in doc.Sections)
            {
                ct.ThrowIfCancellationRequested();

                dynamic container = inHeader ? section.Headers : section.Footers;
                dynamic range = container[ComInteropConstants.WdHeaderFooterPrimary].Range;

                range.Text = string.Empty;
                range.ParagraphFormat.Alignment = wdAlignment;

                if (includeTotalPages)
                {
                    range.InsertAfter("Page ");
                    range.Collapse(ComInteropConstants.WdCollapseEnd);
                    doc.Fields.Add(range, ComInteropConstants.WdFieldPage);

                    dynamic after = container[ComInteropConstants.WdHeaderFooterPrimary].Range;
                    after.Collapse(ComInteropConstants.WdCollapseEnd);
                    after.InsertAfter(" of ");
                    after.Collapse(ComInteropConstants.WdCollapseEnd);
                    doc.Fields.Add(after, ComInteropConstants.WdFieldNumPages);
                }
                else
                {
                    doc.Fields.Add(range, ComInteropConstants.WdFieldPage);
                }

                inserted++;
            }

            string where = inHeader ? "header" : "footer";

            return new FieldResult
            {
                UpdatedCount = inserted,
                Message = $"Page number inserted into the {where} of {inserted} section(s)."
            };
        });
    }

    private static bool ParsePosition(string position)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(position);

        return position.Trim().ToLowerInvariant() switch
        {
            "footer" or "bottom" => false,
            "header" or "top" => true,
            _ => throw new ArgumentException(
                $"Unknown position '{position}'. Use 'header' or 'footer'.", nameof(position))
        };
    }

    private static dynamic ResolveInsertRange(dynamic doc, int? paragraphIndex)
    {
        if (paragraphIndex is not int index)
        {
            dynamic start = doc.Content;
            start.Collapse(ComInteropConstants.WdCollapseStart);
            return start;
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
    /// Counts the entries of a table of contents.
    /// </summary>
    /// <remarks>
    /// Word has no entry count, so the paragraphs of the field result are counted. An empty table
    /// of contents still holds one placeholder paragraph, which is why blank lines are skipped.
    /// </remarks>
    /// <summary>
    /// Counts the entries of a table of contents. Word exposes no entry count, and reading the
    /// generated text does not work: an empty table still holds a placeholder paragraph, and every
    /// paragraph inside the table reports a field because the whole table *is* one TOC field.
    /// Counting the source paragraphs instead is exact and language independent, because
    /// <c>OutlineLevel</c> carries the heading level regardless of the localised style name.
    /// </summary>
    private static int CountEntries(dynamic doc, dynamic toc)
    {
        try
        {
            int upper = (int)toc.UpperHeadingLevel;
            int lower = (int)toc.LowerHeadingLevel;

            var tocRanges = new List<(int Start, int End)>();
            foreach (dynamic other in doc.TablesOfContents)
            {
                dynamic otherRange = other.Range;
                tocRanges.Add(((int)otherRange.Start, (int)otherRange.End));
            }

            int count = 0;
            foreach (dynamic paragraph in doc.Paragraphs)
            {
                int level = (int)paragraph.OutlineLevel;
                if (level < upper || level > lower)
                {
                    continue;
                }

                // Skip anything living inside a table of contents; the generated entries keep the
                // outline level of the heading they point at.
                int start = (int)paragraph.Range.Start;
                if (tocRanges.Exists(r => start >= r.Start && start < r.End))
                {
                    continue;
                }

                count++;
            }

            return count;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            return 0;
        }
    }

    private static int UpdateHeaderFooterFields(dynamic container)
    {
        int updated = 0;

        foreach (dynamic headerFooter in container)
        {
            dynamic fields = headerFooter.Range.Fields;
            int count = (int)fields.Count;
            if (count == 0)
            {
                continue;
            }

            fields.Update();
            updated += count;
        }

        return updated;
    }

    private static Models.FieldInfo Describe(dynamic field, int index)
    {
        string code;
        try
        {
            code = ((string?)field.Code.Text ?? string.Empty).Trim();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            code = string.Empty;
        }

        string result;
        try
        {
            result = WordConversions.CleanRangeText((string?)field.Result.Text);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            result = string.Empty;
        }

        return new Models.FieldInfo
        {
            Index = index,
            Type = (int)field.Type,
            Code = code,
            Result = result
        };
    }
}

using System.Runtime.InteropServices;
using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;
using Word = Microsoft.Office.Interop.Word;

namespace WordMcp.Core.Commands.List;

/// <summary>
/// Word COM implementation of <see cref="IListCommands"/>.
/// </summary>
public sealed class ListCommands : IListCommands
{
    /// <inheritdoc />
    public ListResult Get(IWordBatch batch, bool listedOnly = true)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic paragraphs = doc.Paragraphs;
            int total = (int)paragraphs.Count;

            var list = new List<ListParagraphInfo>();

            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();

                var info = Describe(paragraphs[i], i);
                if (listedOnly && info.ListType == "none")
                    continue;

                list.Add(info);
            }

            return new ListResult
            {
                Paragraphs = list,
                UpdatedCount = 0,
                TotalCount = total
            };
        });
    }

    /// <inheritdoc />
    public ListResult Apply(
        IWordBatch batch,
        int startIndex,
        int? endIndex = null,
        string listType = "bullet",
        int level = 1,
        bool continuePreviousList = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(startIndex, 1);
        ValidateLevel(level);

        int gallery = WordConversions.ToWdListGallery(listType);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            (int first, int last) = ResolveRange((int)doc.Paragraphs.Count, startIndex, endIndex);

            // ListGalleries lives on the Application, not on the Document. The property is
            // strongly typed, so the gallery constant has to be cast back to its enum.
            dynamic template = ctx.App.ListGalleries[(Word.WdListGalleryType)gallery].ListTemplates[1];
            dynamic range = RangeOf(doc, first, last);

            range.ListFormat.ApplyListTemplateWithLevel(
                template,
                continuePreviousList,
                ComInteropConstants.WdListApplyToSelection,
                ComInteropConstants.WdWord10ListBehavior,
                1);

            // The gallery template always lands on level 1, so a deeper level is a second step.
            if (level > 1)
            {
                SetLevels(doc, first, last, level);
            }

            return Result(doc, first, last, $"List formatting applied to paragraphs {first}-{last}.");
        });
    }

    /// <inheritdoc />
    public ListResult SetLevel(IWordBatch batch, int startIndex, int level, int? endIndex = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(startIndex, 1);
        ValidateLevel(level);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            (int first, int last) = ResolveRange((int)doc.Paragraphs.Count, startIndex, endIndex);

            SetLevels(doc, first, last, level);

            return Result(doc, first, last, $"Paragraphs {first}-{last} set to list level {level}.");
        });
    }

    /// <inheritdoc />
    public ListResult Restart(IWordBatch batch, int startIndex, int? endIndex = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(startIndex, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            (int first, int _) = ResolveRange((int)doc.Paragraphs.Count, startIndex, endIndex);

            dynamic startParagraph = doc.Paragraphs[first];
            dynamic startFormat = startParagraph.Range.ListFormat;

            if ((int)startFormat.ListType == ComInteropConstants.WdListNoNumbering)
            {
                throw new ArgumentException(
                    $"Paragraph {first} is not part of a list, so there is no numbering to restart.",
                    nameof(startIndex));
            }

            // Without an explicit end, the restart has to cover the rest of the list; stopping
            // earlier would split it into a third list at that point.
            int last = endIndex ?? FindListEnd(doc, first);

            dynamic template = startFormat.ListTemplate;
            dynamic range = RangeOf(doc, first, last);

            // Re-applying the same template with ContinuePreviousList off is what makes Word treat
            // this as a new list and start counting at 1 again. ApplyTo has to be the range rather
            // than the whole list, or Word keeps the paragraphs in the original list and the
            // numbering just carries on.
            range.ListFormat.ApplyListTemplateWithLevel(
                template,
                false,
                ComInteropConstants.WdListApplyToSelection,
                ComInteropConstants.WdWord10ListBehavior,
                1);

            return Result(doc, first, last, $"Numbering restarted at paragraph {first}.");
        });
    }

    /// <inheritdoc />
    public ListResult Remove(IWordBatch batch, int startIndex, int? endIndex = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(startIndex, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            (int first, int last) = ResolveRange((int)doc.Paragraphs.Count, startIndex, endIndex);

            RangeOf(doc, first, last).ListFormat.RemoveNumbers(ComInteropConstants.WdNumberAllNumbers);

            return Result(doc, first, last, $"List formatting removed from paragraphs {first}-{last}.");
        });
    }

    private static void ValidateLevel(int level)
    {
        if (level is < 1 or > ComInteropConstants.MaxListLevel)
        {
            throw new ArgumentOutOfRangeException(
                nameof(level), level, $"level must be between 1 and {ComInteropConstants.MaxListLevel}.");
        }
    }

    private static (int First, int Last) ResolveRange(int total, int startIndex, int? endIndex)
    {
        if (startIndex > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(startIndex),
                $"Paragraph {startIndex} does not exist. The document has {total} paragraph(s).");
        }

        int last = endIndex ?? startIndex;

        if (last < startIndex)
        {
            throw new ArgumentOutOfRangeException(
                nameof(endIndex), $"end_index {last} is before start_index {startIndex}.");
        }

        if (last > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(endIndex),
                $"Paragraph {last} does not exist. The document has {total} paragraph(s).");
        }

        return (startIndex, last);
    }

    private static dynamic RangeOf(dynamic doc, int first, int last)
        => doc.Range(doc.Paragraphs[first].Range.Start, doc.Paragraphs[last].Range.End);

    private static void SetLevels(dynamic doc, int first, int last, int level)
    {
        for (int i = first; i <= last; i++)
        {
            dynamic format = doc.Paragraphs[i].Range.ListFormat;

            if ((int)format.ListType == ComInteropConstants.WdListNoNumbering)
            {
                throw new ArgumentException(
                    $"Paragraph {i} is not part of a list. Apply list formatting before setting a level.",
                    nameof(level));
            }

            try
            {
                format.ListLevelNumber = level;
            }
            catch (COMException ex)
            {
                throw new ArgumentException(
                    $"Word rejected list level {level} for paragraph {i}. Only outline-number lists "
                        + "support more than one level.",
                    nameof(level),
                    ex);
            }
        }
    }

    /// <summary>
    /// Walks forward from a paragraph to the last paragraph that still belongs to the same list.
    /// </summary>
    private static int FindListEnd(dynamic doc, int first)
    {
        int total = (int)doc.Paragraphs.Count;
        int last = first;

        for (int i = first + 1; i <= total; i++)
        {
            if ((int)doc.Paragraphs[i].Range.ListFormat.ListType == ComInteropConstants.WdListNoNumbering)
                break;

            last = i;
        }

        return last;
    }

    private static ListResult Result(dynamic doc, int first, int last, string message)
    {
        var touched = new List<ListParagraphInfo>();

        for (int i = first; i <= last; i++)
        {
            touched.Add(Describe(doc.Paragraphs[i], i));
        }

        return new ListResult
        {
            Paragraphs = touched,
            UpdatedCount = touched.Count,
            TotalCount = (int)doc.Paragraphs.Count,
            Message = message
        };
    }

    private static ListParagraphInfo Describe(dynamic paragraph, int index)
    {
        dynamic range = paragraph.Range;
        dynamic format = range.ListFormat;

        int listType = (int)format.ListType;
        bool inList = listType != ComInteropConstants.WdListNoNumbering;

        return new ListParagraphInfo
        {
            Index = index,
            Text = WordConversions.CleanRangeText((string?)range.Text),
            ListType = WordConversions.FromWdListType(listType),
            Level = inList ? ReadLevel(format) : 0,
            ListLabel = inList ? ((string?)format.ListString ?? string.Empty).Trim() : string.Empty
        };
    }

    private static int ReadLevel(dynamic format)
    {
        try
        {
            return (int)format.ListLevelNumber;
        }
        catch (COMException)
        {
            // Word refuses to report a level for some list shapes; the list type still is useful.
            return 0;
        }
    }
}

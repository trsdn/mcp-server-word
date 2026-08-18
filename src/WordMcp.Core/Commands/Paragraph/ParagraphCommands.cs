using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Paragraph;

/// <summary>
/// Word COM implementation of <see cref="IParagraphCommands"/>.
/// </summary>
public sealed class ParagraphCommands : IParagraphCommands
{
    /// <inheritdoc />
    public ParagraphListResult List(IWordBatch batch, int skip = 0, int take = 200, bool includeEmpty = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfNegative(skip);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(take);

        return batch.Execute((ctx, ct) =>
        {
            dynamic paragraphs = ctx.Document.Paragraphs;
            int total = (int)paragraphs.Count;

            var list = new List<ParagraphInfo>();
            int emitted = 0;
            int matched = 0;

            for (int i = 1; i <= total && emitted < take; i++)
            {
                ct.ThrowIfCancellationRequested();

                dynamic paragraph = paragraphs[i];
                string text = WordConversions.CleanRangeText((string?)paragraph.Range.Text);

                if (!includeEmpty && string.IsNullOrWhiteSpace(text))
                    continue;

                matched++;
                if (matched <= skip)
                    continue;

                list.Add(Describe(paragraph, i, text));
                emitted++;
            }

            return new ParagraphListResult
            {
                TotalCount = total,
                Paragraphs = list
            };
        });
    }

    /// <inheritdoc />
    public ParagraphResult Add(IWordBatch batch, string text, string? style = null, string? alignment = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        int? wdAlignment = alignment is null ? null : WordConversions.ToWdAlignment(alignment);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic range = doc.Content;
            range.Collapse(ComInteropConstants.WdCollapseEnd);
            range.InsertParagraphAfter();

            int index = (int)doc.Paragraphs.Count;
            dynamic paragraph = doc.Paragraphs[index];
            paragraph.Range.Text = text;

            ApplyStyle(paragraph, style);
            ApplyAlignment(paragraph, wdAlignment);

            return new ParagraphResult
            {
                Paragraph = Describe(paragraph, index, text),
                TotalCount = (int)doc.Paragraphs.Count,
                Message = $"Paragraph added at index {index}."
            };
        });
    }

    /// <inheritdoc />
    public ParagraphResult Insert(IWordBatch batch, int index, string text, string? style = null, string? alignment = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        int? wdAlignment = alignment is null ? null : WordConversions.ToWdAlignment(alignment);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int total = (int)doc.Paragraphs.Count;
            if (index > total)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(index), $"Paragraph {index} does not exist. The document has {total} paragraph(s).");
            }

            dynamic range = doc.Paragraphs[index].Range;
            range.Collapse(ComInteropConstants.WdCollapseStart);
            range.InsertParagraphBefore();

            dynamic inserted = doc.Paragraphs[index];
            inserted.Range.Text = text;

            ApplyStyle(inserted, style);
            ApplyAlignment(inserted, wdAlignment);

            return new ParagraphResult
            {
                Paragraph = Describe(inserted, index, text),
                TotalCount = (int)doc.Paragraphs.Count,
                Message = $"Paragraph inserted at index {index}."
            };
        });
    }

    /// <inheritdoc />
    public ParagraphResult Delete(IWordBatch batch, int index)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int total = (int)doc.Paragraphs.Count;
            if (index > total)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(index), $"Paragraph {index} does not exist. The document has {total} paragraph(s).");
            }

            doc.Paragraphs[index].Range.Delete();

            return new ParagraphResult
            {
                TotalCount = (int)doc.Paragraphs.Count,
                Message = $"Paragraph {index} deleted."
            };
        });
    }

    /// <inheritdoc />
    public ParagraphResult SetStyle(IWordBatch batch, int index, string style)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(style);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return Update(batch, index, paragraph => ApplyStyle(paragraph, style), $"Style '{style}' applied.");
    }

    /// <inheritdoc />
    public ParagraphResult SetAlignment(IWordBatch batch, int index, string alignment)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(alignment);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        int wdAlignment = WordConversions.ToWdAlignment(alignment);
        return Update(batch, index, paragraph => ApplyAlignment(paragraph, wdAlignment), $"Alignment set to {alignment}.");
    }

    private static ParagraphResult Update(IWordBatch batch, int index, Action<dynamic> mutate, string message)
        => batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int total = (int)doc.Paragraphs.Count;
            if (index > total)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(index), $"Paragraph {index} does not exist. The document has {total} paragraph(s).");
            }

            dynamic paragraph = doc.Paragraphs[index];
            mutate(paragraph);

            string text = WordConversions.CleanRangeText((string?)paragraph.Range.Text);

            return new ParagraphResult
            {
                Paragraph = Describe(paragraph, index, text),
                TotalCount = total,
                Message = message
            };
        });

    private static void ApplyStyle(dynamic paragraph, string? style)
    {
        if (string.IsNullOrWhiteSpace(style))
            return;

        try
        {
            paragraph.Range.Style = style;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            throw new ArgumentException(
                $"Style '{style}' does not exist in this document.", nameof(style), ex);
        }
    }

    private static void ApplyAlignment(dynamic paragraph, int? wdAlignment)
    {
        if (wdAlignment.HasValue)
        {
            paragraph.Alignment = wdAlignment.Value;
        }
    }

    private static ParagraphInfo Describe(dynamic paragraph, int index, string text)
    {
        string styleName;
        try
        {
            dynamic style = paragraph.Range.Style;
            styleName = (string)style.NameLocal;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            styleName = string.Empty;
        }

        return new ParagraphInfo
        {
            Index = index,
            Text = text,
            Style = styleName,
            Alignment = WordConversions.FromWdAlignment((int)paragraph.Alignment),
            OutlineLevel = (int)paragraph.OutlineLevel
        };
    }
}

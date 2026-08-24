using System.Runtime.InteropServices;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Footnote;

/// <summary>
/// Word COM implementation of <see cref="IFootnoteCommands"/>.
/// </summary>
public sealed class FootnoteCommands : IFootnoteCommands
{
    /// <inheritdoc />
    public FootnoteListResult List(IWordBatch batch, string kind = "footnote", int maxTextLength = 500)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(maxTextLength, 1);

        string normalized = NormalizeKind(kind);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic notes = Collection(doc, normalized);
            int total = (int)notes.Count;

            var list = new List<FootnoteInfo>();

            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();
                list.Add(Describe(doc, notes[i], i, normalized, maxTextLength));
            }

            return new FootnoteListResult
            {
                Kind = normalized,
                TotalCount = total,
                Notes = list
            };
        });
    }

    /// <inheritdoc />
    public FootnoteResult Add(
        IWordBatch batch,
        int paragraphIndex,
        string text,
        string kind = "footnote",
        string? anchorText = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(paragraphIndex, 1);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);

        string normalized = NormalizeKind(kind);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int paragraphs = (int)doc.Paragraphs.Count;

            if (paragraphIndex > paragraphs)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(paragraphIndex),
                    $"Paragraph {paragraphIndex} does not exist. The document has {paragraphs} paragraph(s).");
            }

            dynamic range = Anchor(doc, paragraphIndex, anchorText);
            dynamic notes = Collection(doc, normalized);

            // The second argument is a custom reference mark; omitting it keeps Word's automatic
            // numbering, which is what a caller who passes no mark expects.
            dynamic note = notes.Add(range, Type.Missing, text);

            int index = (int)note.Index;

            return new FootnoteResult
            {
                Note = Describe(doc, note, index, normalized, 500),
                TotalCount = (int)Collection(doc, normalized).Count,
                Message = $"Added {normalized} {index} to paragraph {paragraphIndex}."
            };
        });
    }

    /// <inheritdoc />
    public FootnoteResult SetText(IWordBatch batch, int index, string text, string kind = "footnote")
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
        ArgumentNullException.ThrowIfNull(text);

        string normalized = NormalizeKind(kind);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic note = Require(doc, normalized, index);

            note.Range.Text = text;

            return new FootnoteResult
            {
                Note = Describe(doc, Collection(doc, normalized)[index], index, normalized, 500),
                TotalCount = (int)Collection(doc, normalized).Count,
                Message = $"Rewrote {normalized} {index}."
            };
        });
    }

    /// <inheritdoc />
    public FootnoteResult Delete(IWordBatch batch, int index, string kind = "footnote")
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        string normalized = NormalizeKind(kind);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic note = Require(doc, normalized, index);

            var info = Describe(doc, note, index, normalized, 500);
            note.Delete();

            int left = (int)Collection(doc, normalized).Count;

            return new FootnoteResult
            {
                Note = info,
                TotalCount = left,
                Message = left == 0
                    ? $"Deleted {normalized} {index}."
                    : $"Deleted {normalized} {index}; Word renumbered the remaining {left}."
            };
        });
    }

    /// <summary>
    /// Footnotes and endnotes are separate collections in Word, so the kind decides which one every
    /// action works on.
    /// </summary>
    private static string NormalizeKind(string kind)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(kind);

        return kind.Trim().ToLowerInvariant() switch
        {
            "footnote" or "footnotes" => "footnote",
            "endnote" or "endnotes" => "endnote",
            _ => throw new ArgumentException(
                $"Unknown note kind '{kind}'. Use one of: footnote, endnote.", nameof(kind))
        };
    }

    private static dynamic Collection(dynamic doc, string kind)
        => kind == "endnote" ? doc.Endnotes : doc.Footnotes;

    private static dynamic Require(dynamic doc, string kind, int index)
    {
        dynamic notes = Collection(doc, kind);
        int total = (int)notes.Count;

        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index),
                total == 0
                    ? $"The document has no {kind}s."
                    : $"{kind} {index} does not exist. The document has {total}.");
        }

        return notes[index];
    }

    /// <summary>
    /// Resolves where the reference mark goes: after a phrase, or at the end of the paragraph text.
    /// </summary>
    private static dynamic Anchor(dynamic doc, int paragraphIndex, string? anchorText)
    {
        dynamic paragraph = doc.Paragraphs[paragraphIndex].Range;

        if (!string.IsNullOrEmpty(anchorText))
        {
            string paragraphText = WordConversions.CleanRangeText((string?)paragraph.Text);
            int offset = paragraphText.IndexOf(anchorText, StringComparison.Ordinal);

            if (offset < 0)
            {
                throw new ArgumentException(
                    $"Paragraph {paragraphIndex} does not contain '{anchorText}'.", nameof(anchorText));
            }

            int end = (int)paragraph.Start + offset + anchorText.Length;
            return doc.Range(end, end);
        }

        // A paragraph range ends after its mark. Anchoring there would push the reference into the
        // next paragraph, so it is placed just before the mark instead.
        int start = (int)paragraph.Start;
        int position = Math.Max(start, (int)paragraph.End - 1);

        return doc.Range(position, position);
    }

    private static FootnoteInfo Describe(dynamic doc, dynamic note, int index, string kind, int maxTextLength)
    {
        string text = WordConversions.CleanRangeText((string?)note.Range.Text);

        return new FootnoteInfo
        {
            Index = index,
            Kind = kind,
            Reference = ReferenceMark(note),
            ParagraphIndex = ParagraphIndexOf(doc, note),
            Text = text.Length > maxTextLength ? text[..maxTextLength] : text
        };
    }

    private static string ReferenceMark(dynamic note)
    {
        try
        {
            return WordConversions.CleanRangeText((string?)note.Reference.Text);
        }
        catch (COMException)
        {
            return string.Empty;
        }
    }

    private static int ParagraphIndexOf(dynamic doc, dynamic note)
    {
        try
        {
            int start = (int)note.Reference.Start;
            int count = (int)doc.Paragraphs.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic range = doc.Paragraphs[i].Range;
                if (start >= (int)range.Start && start < (int)range.End)
                    return i;
            }
        }
        catch (COMException)
        {
            // Falling back to 0 keeps the rest of the note usable.
        }

        return 0;
    }
}

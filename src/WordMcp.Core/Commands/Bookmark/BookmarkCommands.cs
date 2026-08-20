using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Bookmark;

/// <summary>
/// Word COM implementation of <see cref="IBookmarkCommands"/>.
/// </summary>
public sealed partial class BookmarkCommands : IBookmarkCommands
{
    /// <inheritdoc />
    public BookmarkListResult List(IWordBatch batch, int maxTextLength = 200)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(maxTextLength, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic bookmarks = doc.Bookmarks;
            int total = (int)bookmarks.Count;

            var list = new List<BookmarkInfo>();

            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();
                list.Add(Describe(doc, bookmarks[i], maxTextLength));
            }

            return new BookmarkListResult
            {
                TotalCount = total,
                Bookmarks = list
            };
        });
    }

    /// <inheritdoc />
    public BookmarkResult Add(
        IWordBatch batch,
        string name,
        int paragraphIndex,
        int? endParagraphIndex = null,
        string? anchorText = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(paragraphIndex, 1);

        ValidateName(name);

        if (endParagraphIndex is int end)
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(end, paragraphIndex);

            if (anchorText is not null)
            {
                throw new ArgumentException(
                    "anchor_text marks a phrase inside a single paragraph and cannot be combined "
                    + "with end_paragraph_index.",
                    nameof(anchorText));
            }
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int paragraphs = (int)doc.Paragraphs.Count;
            int last = endParagraphIndex ?? paragraphIndex;

            if (last > paragraphs)
            {
                throw new ArgumentOutOfRangeException(
                    endParagraphIndex.HasValue ? nameof(endParagraphIndex) : nameof(paragraphIndex),
                    $"Paragraph {last} does not exist. The document has {paragraphs} paragraph(s).");
            }

            if (BookmarkExists(doc, name))
            {
                throw new ArgumentException(
                    $"Bookmark '{name}' already exists. Delete it first or pick another name.",
                    nameof(name));
            }

            dynamic range = Anchor(doc, paragraphIndex, last, anchorText);
            doc.Bookmarks.Add(name, range);

            return new BookmarkResult
            {
                Bookmark = Describe(doc, doc.Bookmarks[name], 200),
                TotalCount = (int)doc.Bookmarks.Count,
                Message = $"Bookmark '{name}' added at paragraph {paragraphIndex}."
            };
        });
    }

    /// <inheritdoc />
    public BookmarkTextResult GetText(IWordBatch batch, string name)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ValidateName(name);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic bookmark = Require(doc, name);

            string text = WordConversions.CleanRangeText((string?)bookmark.Range.Text);

            return new BookmarkTextResult
            {
                Name = (string?)bookmark.Name ?? name,
                Text = text,
                Length = text.Length
            };
        });
    }

    /// <inheritdoc />
    public BookmarkResult Delete(IWordBatch batch, string name)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ValidateName(name);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic bookmark = Require(doc, name);

            var info = Describe(doc, bookmark, 200);
            bookmark.Delete();

            return new BookmarkResult
            {
                Bookmark = info,
                TotalCount = (int)doc.Bookmarks.Count,
                Message = $"Bookmark '{name}' deleted."
            };
        });
    }

    /// <summary>
    /// Word only accepts bookmark names that start with a letter and consist of letters, digits and
    /// underscores. Anything else fails inside Word with a generic "command failed" error, so the
    /// rule is enforced here where a useful message can still be produced.
    /// </summary>
    private static void ValidateName(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        if (name.Length > 40)
        {
            throw new ArgumentException(
                $"Bookmark name '{name}' is longer than the 40 characters Word allows.", nameof(name));
        }

        if (!NamePattern().IsMatch(name))
        {
            throw new ArgumentException(
                $"Bookmark name '{name}' is invalid. Names must start with a letter and may only "
                + "contain letters, digits and underscores - no spaces or punctuation.",
                nameof(name));
        }
    }

    [GeneratedRegex(@"^[A-Za-z][A-Za-z0-9_]*$")]
    private static partial Regex NamePattern();

    private static bool BookmarkExists(dynamic doc, string name)
    {
        try
        {
            return (bool)doc.Bookmarks.Exists(name);
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static dynamic Require(dynamic doc, string name)
    {
        if (!BookmarkExists(doc, name))
        {
            throw new ArgumentOutOfRangeException(
                nameof(name), $"Bookmark '{name}' does not exist in this document.");
        }

        return doc.Bookmarks[name];
    }

    /// <summary>
    /// Resolves the range a bookmark covers: one paragraph, a paragraph range, or a phrase inside a
    /// single paragraph.
    /// </summary>
    private static dynamic Anchor(dynamic doc, int paragraphIndex, int endParagraphIndex, string? anchorText)
    {
        dynamic first = doc.Paragraphs[paragraphIndex].Range;

        if (!string.IsNullOrEmpty(anchorText))
        {
            string paragraphText = WordConversions.CleanRangeText((string?)first.Text);
            int offset = paragraphText.IndexOf(anchorText, StringComparison.Ordinal);

            if (offset < 0)
            {
                throw new ArgumentException(
                    $"Paragraph {paragraphIndex} does not contain '{anchorText}'.", nameof(anchorText));
            }

            int anchorStart = (int)first.Start + offset;
            return doc.Range(anchorStart, anchorStart + anchorText.Length);
        }

        // A paragraph range ends after its mark; dropping it keeps the bookmark on the text itself,
        // which is what get-text is expected to return.
        dynamic lastRange = doc.Paragraphs[endParagraphIndex].Range;
        int start = (int)first.Start;
        int end = Math.Max(start, (int)lastRange.End - 1);

        return doc.Range(start, end);
    }

    private static BookmarkInfo Describe(dynamic doc, dynamic bookmark, int maxTextLength)
    {
        dynamic range = bookmark.Range;
        string text = WordConversions.CleanRangeText((string?)range.Text);

        return new BookmarkInfo
        {
            Name = (string?)bookmark.Name ?? string.Empty,
            Start = (int)range.Start,
            End = (int)range.End,
            Empty = (bool)bookmark.Empty,
            ParagraphIndex = ParagraphIndexOf(doc, (int)range.Start),
            Text = text.Length > maxTextLength ? text[..maxTextLength] : text
        };
    }

    private static int ParagraphIndexOf(dynamic doc, int start)
    {
        try
        {
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
            // Falling back to 0 keeps the rest of the bookmark usable.
        }

        return 0;
    }
}

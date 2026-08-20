using System.Runtime.InteropServices;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Comment;

/// <summary>
/// Word COM implementation of <see cref="ICommentCommands"/>.
/// </summary>
public sealed class CommentCommands : ICommentCommands
{
    /// <inheritdoc />
    public CommentListResult List(IWordBatch batch, bool unresolvedOnly = false)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic comments = doc.Comments;
            int total = (int)comments.Count;

            var list = new List<CommentInfo>();

            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();

                var info = Describe(comments[i], i);
                if (unresolvedOnly && info.Resolved == true)
                    continue;

                list.Add(info);
            }

            return new CommentListResult
            {
                TotalCount = total,
                Comments = list
            };
        });
    }

    /// <inheritdoc />
    public CommentResult Add(IWordBatch batch, int paragraphIndex, string text, string? anchorText = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        ArgumentOutOfRangeException.ThrowIfLessThan(paragraphIndex, 1);

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

            dynamic paragraph = doc.Paragraphs[paragraphIndex];
            dynamic range = Anchor(doc, paragraph, paragraphIndex, anchorText);

            dynamic comment = doc.Comments.Add(range, text);
            int index = (int)comment.Index;

            return new CommentResult
            {
                Comment = Describe(comment, index),
                TotalCount = (int)doc.Comments.Count,
                Message = $"Comment {index} added to paragraph {paragraphIndex}."
            };
        });
    }

    /// <inheritdoc />
    public CommentResult Resolve(IWordBatch batch, int index, bool resolved = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic comment = Require(doc, index);

            try
            {
                comment.Done = resolved;
            }
            catch (COMException ex)
            {
                // Two cases end up here and both surface as "command is not available": Word before
                // 2013 has no Done property at all, and modern comments refuse it for a comment that
                // was never posted in the UI - which includes every comment added through this API.
                throw new NotSupportedException(
                    $"Word refused to change the resolved state of comment {index}. Comments added " +
                    "through the API are drafts that modern comments cannot resolve, and Word " +
                    "versions before 2013 do not support resolving at all. Delete the comment instead.",
                    ex);
            }

            return new CommentResult
            {
                Comment = Describe(comment, index),
                TotalCount = (int)doc.Comments.Count,
                Message = resolved ? $"Comment {index} resolved." : $"Comment {index} reopened."
            };
        });
    }

    /// <inheritdoc />
    public CommentResult Delete(IWordBatch batch, int index)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            Require(doc, index).Delete();

            return new CommentResult
            {
                TotalCount = (int)doc.Comments.Count,
                Message = $"Comment {index} deleted."
            };
        });
    }

    private static dynamic Require(dynamic doc, int index)
    {
        int total = (int)doc.Comments.Count;

        if (index > total)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index), $"Comment {index} does not exist. The document has {total} comment(s).");
        }

        return doc.Comments[index];
    }

    /// <summary>
    /// Resolves the range a comment attaches to: the whole paragraph, or the phrase inside it.
    /// </summary>
    private static dynamic Anchor(dynamic doc, dynamic paragraph, int paragraphIndex, string? anchorText)
    {
        if (string.IsNullOrEmpty(anchorText))
        {
            return paragraph.Range;
        }

        string paragraphText = WordConversions.CleanRangeText((string?)paragraph.Range.Text);
        int offset = paragraphText.IndexOf(anchorText, StringComparison.Ordinal);

        if (offset < 0)
        {
            throw new ArgumentException(
                $"Paragraph {paragraphIndex} does not contain '{anchorText}'.", nameof(anchorText));
        }

        int start = (int)paragraph.Range.Start + offset;
        return doc.Range(start, start + anchorText.Length);
    }

    private static CommentInfo Describe(dynamic comment, int index)
    {
        var info = new CommentInfo
        {
            Index = index,
            Author = (string?)comment.Author ?? string.Empty,
            Initial = (string?)comment.Initial ?? string.Empty,
            Text = WordConversions.CleanRangeText((string?)comment.Range.Text),
            Date = ReadDate(comment),
            Resolved = ReadResolved(comment)
        };

        try
        {
            dynamic scope = comment.Scope;
            info.ScopeText = WordConversions.CleanRangeText((string?)scope.Text);
            info.ParagraphIndex = ParagraphIndexOf(comment, scope);
        }
        catch (COMException)
        {
            // A comment whose anchor was deleted still exists but has no usable scope.
        }

        return info;
    }

    private static int ParagraphIndexOf(dynamic comment, dynamic scope)
    {
        try
        {
            dynamic doc = comment.Parent;
            int start = (int)scope.Start;
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
            // Falling back to 0 keeps the rest of the comment usable.
        }

        return 0;
    }

    private static DateTime? ReadDate(dynamic comment)
    {
        try
        {
            return (DateTime)comment.Date;
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static bool? ReadResolved(dynamic comment)
    {
        try
        {
            return (bool)comment.Done;
        }
        catch (COMException)
        {
            // Word versions before 2013 have no resolved state at all.
            return null;
        }
    }
}

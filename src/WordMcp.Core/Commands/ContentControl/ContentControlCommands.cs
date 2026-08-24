using System.Runtime.InteropServices;
using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.ContentControl;

/// <summary>
/// Word COM implementation of <see cref="IContentControlCommands"/>.
/// </summary>
public sealed class ContentControlCommands : IContentControlCommands
{
    /// <inheritdoc />
    public ContentControlListResult List(IWordBatch batch, string? tag = null, int maxTextLength = 500)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(maxTextLength, 1);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic controls = doc.ContentControls;
            int total = (int)controls.Count;

            var list = new List<ContentControlInfo>();

            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();
                var info = Describe(controls[i], i, maxTextLength);

                if (tag is null || string.Equals(info.Tag, tag, StringComparison.Ordinal))
                    list.Add(info);
            }

            return new ContentControlListResult
            {
                TotalCount = list.Count,
                Controls = list,
                Message = tag is null
                    ? $"Found {list.Count} content control(s)."
                    : $"Found {list.Count} content control(s) tagged '{tag}' out of {total}."
            };
        });
    }

    /// <inheritdoc />
    public ContentControlResult Get(
        IWordBatch batch,
        string? tag = null,
        string? id = null,
        int maxTextLength = 5000)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(maxTextLength, 1);
        RequireSelector(tag, id);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            List<ControlMatch> matches = Match(doc, tag, id, maxTextLength);

            return new ContentControlResult
            {
                Controls = matches.Select(m => m.Info).ToList(),
                AffectedCount = matches.Count,
                Message = $"Read {matches.Count} content control(s)."
            };
        });
    }

    /// <inheritdoc />
    public ContentControlResult SetText(
        IWordBatch batch,
        string? text = null,
        string? tag = null,
        string? id = null,
        bool? isChecked = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        RequireSelector(tag, id);

        if (text is null && isChecked is null)
        {
            throw new ArgumentException(
                "Provide text for a text control, or is_checked for a checkbox control.", nameof(text));
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            List<ControlMatch> matches = Match(doc, tag, id, 5000);
            var updated = new List<ContentControlInfo>();

            foreach (var match in matches)
            {
                ct.ThrowIfCancellationRequested();
                dynamic control = match.Control;

                if (match.Info.LockContents)
                {
                    throw new InvalidOperationException(
                        $"Content control '{match.Info.Tag}' (id {match.Info.Id}) has locked contents. "
                        + "Unlock it in Word before writing to it.");
                }

                int type = (int)control.Type;

                if (type == ComInteropConstants.WdContentControlCheckBox)
                {
                    if (isChecked is null)
                    {
                        throw new ArgumentException(
                            $"Content control '{match.Info.Tag}' is a checkbox; pass is_checked instead of text.",
                            nameof(isChecked));
                    }

                    control.Checked = isChecked.Value;
                }
                else
                {
                    if (text is null)
                    {
                        throw new ArgumentException(
                            $"Content control '{match.Info.Tag}' is a {match.Info.Type} control; pass text.",
                            nameof(text));
                    }

                    SetControlText(control, type, text, match.Info);
                }

                updated.Add(Describe(control, match.Info.Index, 5000));
            }

            return new ContentControlResult
            {
                Controls = updated,
                AffectedCount = updated.Count,
                Message = updated.Count == 1
                    ? "Updated 1 content control."
                    : $"Updated {updated.Count} content controls."
            };
        });
    }

    /// <inheritdoc />
    public ContentControlResult Add(
        IWordBatch batch,
        int paragraphIndex,
        string type = "rich-text",
        string? tag = null,
        string? title = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfLessThan(paragraphIndex, 1);

        int wdType = WordConversions.ToWdContentControlType(type);

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

            dynamic paragraph = doc.Paragraphs[paragraphIndex].Range;

            // A paragraph range ends after its mark; including it would make Word wrap the following
            // paragraph as well.
            int start = (int)paragraph.Start;
            int end = Math.Max(start, (int)paragraph.End - 1);
            dynamic range = doc.Range(start, end);

            dynamic control = doc.ContentControls.Add(wdType, range);

            if (!string.IsNullOrEmpty(tag))
                control.Tag = tag;

            if (!string.IsNullOrEmpty(title))
                control.Title = title;

            var info = Describe(control, (int)doc.ContentControls.Count, 5000);

            return new ContentControlResult
            {
                Controls = [info],
                AffectedCount = 1,
                Message = $"Added a {info.Type} content control around paragraph {paragraphIndex}."
            };
        });
    }

    /// <inheritdoc />
    public ContentControlResult Delete(
        IWordBatch batch,
        string? tag = null,
        string? id = null,
        bool deleteContents = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        RequireSelector(tag, id);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            List<ControlMatch> matches = Match(doc, tag, id, 5000);

            foreach (var match in matches)
            {
                ct.ThrowIfCancellationRequested();

                if (match.Info.LockContentControl)
                {
                    throw new InvalidOperationException(
                        $"Content control '{match.Info.Tag}' (id {match.Info.Id}) is locked against "
                        + "deletion. Unlock it in Word first.");
                }

                match.Control.Delete(deleteContents);
            }

            return new ContentControlResult
            {
                Controls = matches.Select(m => m.Info).ToList(),
                AffectedCount = matches.Count,
                Message = deleteContents
                    ? $"Deleted {matches.Count} content control(s) and their contents."
                    : $"Deleted {matches.Count} content control(s); their text stayed in the document."
            };
        });
    }

    private static void RequireSelector(string? tag, string? id)
    {
        if (string.IsNullOrWhiteSpace(tag) && string.IsNullOrWhiteSpace(id))
        {
            throw new ArgumentException(
                "Provide either tag or id to address a content control. Call content-control(list) "
                + "to see what the document offers.", nameof(tag));
        }
    }

    /// <summary>
    /// Resolves the controls a caller addressed. Tags are deliberately allowed to match several
    /// controls, because templates repeat a field on purpose.
    /// </summary>
    private static List<ControlMatch> Match(dynamic doc, string? tag, string? id, int maxTextLength)
    {
        dynamic controls = doc.ContentControls;
        int total = (int)controls.Count;
        var matches = new List<ControlMatch>();

        for (int i = 1; i <= total; i++)
        {
            dynamic control = controls[i];
            var info = Describe(control, i, maxTextLength);

            bool hit = id is not null
                ? string.Equals(info.Id, id, StringComparison.OrdinalIgnoreCase)
                : string.Equals(info.Tag, tag, StringComparison.Ordinal);

            if (hit)
                matches.Add(new ControlMatch(control, info));
        }

        if (matches.Count == 0)
        {
            string what = id is not null ? $"id '{id}'" : $"tag '{tag}'";
            throw new ArgumentException(
                $"No content control with {what} found. The document has {total} content control(s); "
                + "call content-control(list) to see their tags.",
                id is not null ? nameof(id) : nameof(tag));
        }

        return matches;
    }

    /// <summary>
    /// Writes text into a control. Word rejects a direct assignment while the placeholder is showing
    /// for some control types, so the range is cleared first.
    /// </summary>
    private static void SetControlText(dynamic control, int type, string text, ContentControlInfo info)
    {
        if (type is ComInteropConstants.WdContentControlPicture
            or ComInteropConstants.WdContentControlGroup
            or ComInteropConstants.WdContentControlRepeatingSection
            or ComInteropConstants.WdContentControlBuildingBlockGallery)
        {
            throw new InvalidOperationException(
                $"Content control '{info.Tag}' is a {info.Type} control and holds no plain text.");
        }

        try
        {
            control.Range.Text = text;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                $"Word refused to write into content control '{info.Tag}' ({info.Type}): {ex.Message}", ex);
        }
    }

    private static ContentControlInfo Describe(dynamic control, int index, int maxTextLength)
    {
        int type = (int)control.Type;
        string text = WordConversions.CleanRangeText((string?)control.Range.Text);

        return new ContentControlInfo
        {
            Index = index,
            Id = (string?)control.ID ?? string.Empty,
            Tag = (string?)control.Tag ?? string.Empty,
            Title = (string?)control.Title ?? string.Empty,
            Type = WordConversions.FromWdContentControlType(type),
            Text = text.Length > maxTextLength ? text[..maxTextLength] : text,
            LockContents = (bool)control.LockContents,
            LockContentControl = (bool)control.LockContentControl,
            Checked = type == ComInteropConstants.WdContentControlCheckBox ? (bool)control.Checked : null,
            ShowingPlaceholderText = ShowingPlaceholder(control)
        };
    }

    private static bool ShowingPlaceholder(dynamic control)
    {
        try
        {
            return (bool)control.ShowingPlaceholderText;
        }
        catch (COMException)
        {
            return false;
        }
    }

    private sealed record ControlMatch(dynamic Control, ContentControlInfo Info);
}

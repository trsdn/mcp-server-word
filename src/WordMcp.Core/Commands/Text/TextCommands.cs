using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Text;

/// <summary>
/// Word COM implementation of <see cref="ITextCommands"/>.
/// </summary>
public sealed class TextCommands : ITextCommands
{
    private const int ContextRadius = 40;

    /// <inheritdoc />
    public TextResult Get(IWordBatch batch, int? start = null, int? end = null, int maxLength = 100_000)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxLength);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int docEnd = (int)doc.Content.End;

            int rangeStart = Math.Clamp(start ?? 0, 0, docEnd);
            int rangeEnd = Math.Clamp(end ?? docEnd, rangeStart, docEnd);

            dynamic range = doc.Range(rangeStart, rangeEnd);
            string text = WordConversions.CleanRangeText((string?)range.Text);

            bool truncated = text.Length > maxLength;
            if (truncated)
            {
                text = text[..maxLength];
            }

            return new TextResult
            {
                Text = text,
                Start = rangeStart,
                End = rangeEnd,
                Truncated = truncated
            };
        });
    }

    /// <inheritdoc />
    public TextResult Append(IWordBatch batch, string text, bool newParagraph = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic range = doc.Content;
            range.Collapse(ComInteropConstants.WdCollapseEnd);

            if (newParagraph)
            {
                range.InsertParagraphAfter();
                range = doc.Content;
                range.Collapse(ComInteropConstants.WdCollapseEnd);
            }

            int insertStart = (int)range.Start;
            range.InsertAfter(text);

            return new TextResult
            {
                Text = text,
                Start = insertStart,
                End = (int)doc.Content.End,
                Message = "Text appended."
            };
        });
    }

    /// <inheritdoc />
    public FindResult Find(
        IWordBatch batch,
        string searchText,
        bool matchCase = false,
        bool matchWholeWord = false,
        int maxResults = 100)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrEmpty(searchText);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxResults);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            string content = WordConversions.CleanRangeText((string?)doc.Content.Text);
            int contentStart = (int)doc.Content.Start;

            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var matches = new List<TextMatch>();
            int total = 0;
            int index = 0;

            while (index <= content.Length - searchText.Length)
            {
                int hit = content.IndexOf(searchText, index, comparison);
                if (hit < 0)
                    break;

                index = hit + searchText.Length;

                if (matchWholeWord && !IsWholeWord(content, hit, searchText.Length))
                    continue;

                total++;
                if (matches.Count < maxResults)
                {
                    int ctxStart = Math.Max(0, hit - ContextRadius);
                    int ctxEnd = Math.Min(content.Length, hit + searchText.Length + ContextRadius);

                    matches.Add(new TextMatch
                    {
                        Start = contentStart + hit,
                        End = contentStart + hit + searchText.Length,
                        Context = content[ctxStart..ctxEnd].Replace('\r', ' ')
                    });
                }
            }

            return new FindResult
            {
                SearchText = searchText,
                MatchCount = total,
                Matches = matches,
                Message = total > matches.Count
                    ? $"{total} matches found, first {matches.Count} returned."
                    : $"{total} matches found."
            };
        });
    }

    /// <inheritdoc />
    public ReplaceResult Replace(
        IWordBatch batch,
        string searchText,
        string replaceText,
        bool matchCase = false,
        bool matchWholeWord = false,
        bool replaceAll = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrEmpty(searchText);
        ArgumentNullException.ThrowIfNull(replaceText);

        var before = Find(batch, searchText, matchCase, matchWholeWord, int.MaxValue);
        int expected = replaceAll ? before.MatchCount : Math.Min(1, before.MatchCount);

        batch.Execute((ctx, ct) =>
        {
            dynamic find = ctx.Document.Content.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();

            // Positional arguments: FindText, MatchCase, MatchWholeWord, MatchWildcards,
            // MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace.
            find.Execute(
                searchText,
                matchCase,
                matchWholeWord,
                false,
                false,
                false,
                true,
                ComInteropConstants.WdFindStop,
                false,
                replaceText,
                replaceAll ? ComInteropConstants.WdReplaceAll : ComInteropConstants.WdReplaceOne);
        });

        return new ReplaceResult
        {
            SearchText = searchText,
            ReplaceText = replaceText,
            ReplacementCount = expected,
            Message = $"{expected} occurrence(s) replaced."
        };
    }

    /// <inheritdoc />
    public OperationResult Format(
        IWordBatch batch,
        int start,
        int end,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? fontName = null,
        double? fontSize = null,
        string? color = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentOutOfRangeException.ThrowIfNegative(start);

        if (end <= start)
        {
            throw new ArgumentOutOfRangeException(nameof(end), "The end position must be greater than the start position.");
        }

        int? wdColor = color is null ? null : WordConversions.ToWdColor(color);

        batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            int docEnd = (int)doc.Content.End;

            int rangeStart = Math.Clamp(start, 0, docEnd);
            int rangeEnd = Math.Clamp(end, rangeStart, docEnd);

            dynamic font = doc.Range(rangeStart, rangeEnd).Font;

            if (bold.HasValue)
                font.Bold = bold.Value ? 1 : 0;
            if (italic.HasValue)
                font.Italic = italic.Value ? 1 : 0;
            if (underline.HasValue)
                font.Underline = underline.Value ? 1 : 0;
            if (!string.IsNullOrWhiteSpace(fontName))
                font.Name = fontName;
            if (fontSize.HasValue)
                font.Size = (float)fontSize.Value;
            if (wdColor.HasValue)
                font.Color = wdColor.Value;
        });

        return OperationResult.Ok($"Formatting applied to characters {start}-{end}.");
    }

    private static bool IsWholeWord(string content, int index, int length)
    {
        bool leftOk = index == 0 || !char.IsLetterOrDigit(content[index - 1]);
        int after = index + length;
        bool rightOk = after >= content.Length || !char.IsLetterOrDigit(content[after]);
        return leftOk && rightOk;
    }
}

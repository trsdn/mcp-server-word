using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>text</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordTextAction>))]
public enum WordTextAction
{
    /// <summary>Read text.</summary>
    [JsonStringEnumMemberName("get")] Get,

    /// <summary>Append text at the end of the document.</summary>
    [JsonStringEnumMemberName("append")] Append,

    /// <summary>Find occurrences of a term.</summary>
    [JsonStringEnumMemberName("find")] Find,

    /// <summary>Replace occurrences of a term.</summary>
    [JsonStringEnumMemberName("replace")] Replace,

    /// <summary>Apply character formatting to a range.</summary>
    [JsonStringEnumMemberName("format")] Format
}

/// <summary>
/// Text operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordTextTool
{
    /// <summary>
    /// Reads, appends, searches, replaces and formats document text.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="text">Text to append (append) or search term (find, replace).</param>
    /// <param name="replace_text">Replacement text for replace.</param>
    /// <param name="start">Start character position for get and format.</param>
    /// <param name="end">End character position for get and format.</param>
    /// <param name="max_length">Maximum number of characters returned by get.</param>
    /// <param name="match_case">Whether find and replace are case sensitive.</param>
    /// <param name="match_whole_word">Whether find and replace match whole words only.</param>
    /// <param name="replace_all">Whether replace changes all occurrences.</param>
    /// <param name="max_results">Maximum number of matches returned by find.</param>
    /// <param name="new_paragraph">Whether append starts a new paragraph.</param>
    /// <param name="bold">Bold for format.</param>
    /// <param name="italic">Italic for format.</param>
    /// <param name="underline">Underline for format.</param>
    /// <param name="font_name">Font family for format.</param>
    /// <param name="font_size">Font size in points for format.</param>
    /// <param name="color">Font colour in hex notation for format.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "text", Title = "Text Operations")]
    [Description("Text operations on an open document. "
        + "text(get, session_id) returns the full text; add start/end to read a character range. "
        + "text(append, session_id, text='...') appends a paragraph at the end. "
        + "text(find, session_id, text='term') returns match positions and context. "
        + "text(replace, session_id, text='old', replace_text='new') replaces occurrences. "
        + "text(format, session_id, start=0, end=20, bold=true, color='#0078D4') formats a character range. "
        + "Positions are Word character offsets as reported by get and find.")]
    public static string Text(
        WordTextAction action,
        string session_id,
        [DefaultValue(null)] string? text = null,
        [DefaultValue(null)] string? replace_text = null,
        [DefaultValue(null)] int? start = null,
        [DefaultValue(null)] int? end = null,
        [DefaultValue(100000)] int max_length = 100_000,
        [DefaultValue(false)] bool match_case = false,
        [DefaultValue(false)] bool match_whole_word = false,
        [DefaultValue(true)] bool replace_all = true,
        [DefaultValue(100)] int max_results = 100,
        [DefaultValue(true)] bool new_paragraph = true,
        [DefaultValue(null)] bool? bold = null,
        [DefaultValue(null)] bool? italic = null,
        [DefaultValue(null)] bool? underline = null,
        [DefaultValue(null)] string? font_name = null,
        [DefaultValue(null)] double? font_size = null,
        [DefaultValue(null)] string? color = null)
        => WordToolsBase.Execute("text", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordTextAction.Get => WordServices.Texts.Get(batch, start, end, max_length),
                WordTextAction.Append => WordServices.Texts.Append(
                    batch, Require(text, nameof(text)), new_paragraph),
                WordTextAction.Find => WordServices.Texts.Find(
                    batch, Require(text, nameof(text)), match_case, match_whole_word, max_results),
                WordTextAction.Replace => WordServices.Texts.Replace(
                    batch, Require(text, nameof(text)), replace_text ?? string.Empty,
                    match_case, match_whole_word, replace_all),
                WordTextAction.Format => WordServices.Texts.Format(
                    batch,
                    start ?? throw new ArgumentException("start is required for 'format'.", nameof(start)),
                    end ?? throw new ArgumentException("end is required for 'format'.", nameof(end)),
                    bold, italic, underline, font_name, font_size, color),
                _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
            };
        });

    private static string Require(string? value, string name)
        => string.IsNullOrEmpty(value)
            ? throw new ArgumentException($"{name} is required for this action.", name)
            : value;
}

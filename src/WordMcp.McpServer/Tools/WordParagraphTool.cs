using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>paragraph</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordParagraphAction>))]
public enum WordParagraphAction
{
    /// <summary>List paragraphs.</summary>
    [JsonStringEnumMemberName("list")] List,

    /// <summary>Append a paragraph.</summary>
    [JsonStringEnumMemberName("add")] Add,

    /// <summary>Insert a paragraph before an existing one.</summary>
    [JsonStringEnumMemberName("insert")] Insert,

    /// <summary>Delete a paragraph.</summary>
    [JsonStringEnumMemberName("delete")] Delete,

    /// <summary>Apply a paragraph style.</summary>
    [JsonStringEnumMemberName("set-style")] SetStyle,

    /// <summary>Set paragraph alignment.</summary>
    [JsonStringEnumMemberName("set-alignment")] SetAlignment
}

/// <summary>
/// Paragraph operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordParagraphTool
{
    /// <summary>
    /// Lists, adds, inserts, deletes and styles paragraphs.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="index">1-based paragraph index.</param>
    /// <param name="text">Paragraph text for add and insert.</param>
    /// <param name="style">Style name, for example <c>Heading 1</c>.</param>
    /// <param name="alignment">Alignment: left, center, right or justify.</param>
    /// <param name="skip">Number of paragraphs to skip in list.</param>
    /// <param name="take">Maximum number of paragraphs returned by list.</param>
    /// <param name="include_empty">Whether list includes empty paragraphs.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "paragraph", Title = "Paragraph Operations")]
    [Description("Paragraph operations on an open document. "
        + "paragraph(list, session_id, skip=0, take=200) returns index, text, style, alignment and outline level. "
        + "paragraph(add, session_id, text='...', style='Heading 1') appends a paragraph. "
        + "paragraph(insert, session_id, index=3, text='...') inserts before paragraph 3. "
        + "paragraph(delete|set-style|set-alignment, session_id, index=3, ...) edits an existing paragraph. "
        + "Indexes are 1-based and shift after add, insert or delete — re-run list before further edits.")]
    public static string Paragraph(
        WordParagraphAction action,
        string session_id,
        [DefaultValue(null)] int? index = null,
        [DefaultValue(null)] string? text = null,
        [DefaultValue(null)] string? style = null,
        [DefaultValue(null)] string? alignment = null,
        [DefaultValue(0)] int skip = 0,
        [DefaultValue(200)] int take = 200,
        [DefaultValue(true)] bool include_empty = true)
        => WordToolsBase.Execute("paragraph", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordParagraphAction.List => WordServices.Paragraphs.List(batch, skip, take, include_empty),
                WordParagraphAction.Add => WordServices.Paragraphs.Add(
                    batch, text ?? string.Empty, style, alignment),
                WordParagraphAction.Insert => WordServices.Paragraphs.Insert(
                    batch, RequireIndex(index), text ?? string.Empty, style, alignment),
                WordParagraphAction.Delete => WordServices.Paragraphs.Delete(batch, RequireIndex(index)),
                WordParagraphAction.SetStyle => WordServices.Paragraphs.SetStyle(
                    batch, RequireIndex(index), Require(style, nameof(style))),
                WordParagraphAction.SetAlignment => WordServices.Paragraphs.SetAlignment(
                    batch, RequireIndex(index), Require(alignment, nameof(alignment))),
                _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
            };
        });

    private static int RequireIndex(int? index)
        => index ?? throw new ArgumentException("index is required for this action.", nameof(index));

    private static string Require(string? value, string name)
        => string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException($"{name} is required for this action.", name)
            : value;
}

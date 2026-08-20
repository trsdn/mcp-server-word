using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>header-footer</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordHeaderFooterAction>))]
public enum WordHeaderFooterAction
{
    /// <summary>Read headers or footers.</summary>
    [JsonStringEnumMemberName("get")] Get,

    /// <summary>Write a header or footer.</summary>
    [JsonStringEnumMemberName("set")] Set,

    /// <summary>Clear a header or footer.</summary>
    [JsonStringEnumMemberName("clear")] Clear
}

/// <summary>
/// Header and footer operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordHeaderFooterTool
{
    /// <summary>
    /// Reads, writes and clears headers and footers.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="kind">Either <c>header</c> or <c>footer</c>.</param>
    /// <param name="text">The text to write; required for set.</param>
    /// <param name="section_index">1-based section; all sections when omitted.</param>
    /// <param name="type">One of <c>primary</c>, <c>first-page</c> or <c>even-pages</c>.</param>
    /// <param name="alignment">Optional paragraph alignment for the written text.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "header-footer", Title = "Header and Footer Operations")]
    [Description("Header and footer operations on an open document. Headers belong to a SECTION, "
        + "not to the document, so use section(list) first when a document has more than one. "
        + "header-footer(set, session_id, kind='footer', text='Confidential') writes the same text "
        + "to every section; pass section_index to target one. "
        + "type='first-page' or 'even-pages' switches the matching section option on automatically, "
        + "because Word otherwise stores the text without ever showing it. "
        + "Writing to a section whose header is inherited from the previous one breaks that link. "
        + "header-footer(get, session_id, kind='header') reads the text back. "
        + "For page numbers use field(insert-page-number) instead, which inserts live fields.")]
    public static string HeaderFooter(
        WordHeaderFooterAction action,
        string session_id,
        [DefaultValue("header")] string kind = "header",
        [DefaultValue(null)] string? text = null,
        [DefaultValue(null)] int? section_index = null,
        [DefaultValue("primary")] string type = "primary",
        [DefaultValue(null)] string? alignment = null)
        => WordToolsBase.Execute("header-footer", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordHeaderFooterAction.Get => WordServices.HeadersFooters.Get(
                    batch, kind, section_index, type),
                WordHeaderFooterAction.Set => WordServices.HeadersFooters.Set(
                    batch,
                    text ?? throw new ArgumentException(
                        "text is required for header-footer(set).", nameof(text)),
                    kind,
                    section_index,
                    type,
                    alignment),
                WordHeaderFooterAction.Clear => WordServices.HeadersFooters.Clear(
                    batch, kind, section_index, type),
                _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
            };
        });
}

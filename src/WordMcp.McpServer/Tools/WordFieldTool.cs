using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>field</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordFieldAction>))]
public enum WordFieldAction
{
    /// <summary>List all fields.</summary>
    [JsonStringEnumMemberName("list")] List,

    /// <summary>Insert a table of contents.</summary>
    [JsonStringEnumMemberName("insert-toc")] InsertToc,

    /// <summary>Update every table of contents.</summary>
    [JsonStringEnumMemberName("update-toc")] UpdateToc,

    /// <summary>Update all fields.</summary>
    [JsonStringEnumMemberName("update-all")] UpdateAll,

    /// <summary>Insert a page number field.</summary>
    [JsonStringEnumMemberName("insert-page-number")] InsertPageNumber
}

/// <summary>
/// Field operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordFieldTool
{
    /// <summary>
    /// Inserts and updates fields such as a table of contents and page numbers.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="paragraph_index">1-based paragraph to insert before; inserts at the top when omitted.</param>
    /// <param name="upper_heading_level">Highest heading level of the table of contents.</param>
    /// <param name="lower_heading_level">Lowest heading level of the table of contents.</param>
    /// <param name="include_page_numbers">Whether the table of contents shows page numbers.</param>
    /// <param name="use_hyperlinks">Whether table of contents entries link to their heading.</param>
    /// <param name="position">Either <c>footer</c> or <c>header</c> for page numbers.</param>
    /// <param name="alignment">One of <c>left</c>, <c>center</c> or <c>right</c> for page numbers.</param>
    /// <param name="include_total_pages">Whether to render "Page X of Y".</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "field", Title = "Field Operations")]
    [Description("Field operations on an open document. "
        + "field(insert-toc, session_id, lower_heading_level=3) inserts a table of contents at the top "
        + "and fills it immediately. It stays EMPTY unless paragraphs use heading styles, so apply "
        + "'Heading 1'/'Heading 2' via paragraph(add, style=...) first; the result reports entry_count. "
        + "field(insert-page-number, session_id, position='footer', alignment='center', include_total_pages=true) "
        + "adds page numbers to every section. "
        + "field(update-toc, session_id) refreshes the table of contents after content changed. "
        + "field(update-all, session_id) refreshes every field including headers and footers. "
        + "field(list, session_id) returns type, code and current result of each field.")]
    public static string Field(
        WordFieldAction action,
        string session_id,
        [DefaultValue(null)] int? paragraph_index = null,
        [DefaultValue(1)] int upper_heading_level = 1,
        [DefaultValue(3)] int lower_heading_level = 3,
        [DefaultValue(true)] bool include_page_numbers = true,
        [DefaultValue(true)] bool use_hyperlinks = true,
        [DefaultValue("footer")] string position = "footer",
        [DefaultValue("center")] string alignment = "center",
        [DefaultValue(false)] bool include_total_pages = false)
        => WordToolsBase.Execute("field", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordFieldAction.List => WordServices.Fields.List(batch),
                WordFieldAction.InsertToc => WordServices.Fields.InsertTableOfContents(
                    batch,
                    paragraph_index,
                    upper_heading_level,
                    lower_heading_level,
                    include_page_numbers,
                    use_hyperlinks),
                WordFieldAction.UpdateToc => WordServices.Fields.UpdateTableOfContents(batch),
                WordFieldAction.UpdateAll => WordServices.Fields.UpdateAll(batch),
                WordFieldAction.InsertPageNumber => WordServices.Fields.InsertPageNumber(
                    batch, position, alignment, include_total_pages),
                _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
            };
        });
}

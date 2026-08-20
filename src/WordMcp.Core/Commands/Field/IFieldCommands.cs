using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Field;

/// <summary>
/// Field operations: table of contents, page numbers and field updates.
/// </summary>
/// <remarks>
/// Fields hold a code and a cached result. The result is only recalculated on update, so a freshly
/// inserted table of contents stays empty until <c>update-toc</c> or <c>update-all</c> runs — the
/// insert actions therefore update right away.
/// </remarks>
[ServiceCategory("field", "Field")]
[McpTool("field",
    Title = "Field Operations",
    Description = "Field operations on an open document. "
        + "field(insert-toc, session_id, lower_heading_level=3) inserts a table of contents at the top "
        + "and fills it immediately. It stays EMPTY unless paragraphs use heading styles, so apply "
        + "'Heading 1'/'Heading 2' via paragraph(add, style=...) first; the result reports entry_count. "
        + "field(insert-page-number, session_id, position='footer', alignment='center', include_total_pages=true) "
        + "adds page numbers to every section. "
        + "field(update-toc, session_id) refreshes the table of contents after content changed. "
        + "field(update-all, session_id) refreshes every field including headers and footers. "
        + "field(list, session_id) returns type, code and current result of each field.")]
public interface IFieldCommands
{
    /// <summary>
    /// Lists all fields of the main document text.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>The fields of the document.</returns>
    [ServiceAction("list")]
    FieldListResult List(IWordBatch batch);

    /// <summary>
    /// Inserts a table of contents and updates it immediately.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="paragraphIndex">
    /// 1-based paragraph to insert before. When omitted the table of contents goes to the top.
    /// </param>
    /// <param name="upperHeadingLevel">Highest heading level to include. Defaults to 1.</param>
    /// <param name="lowerHeadingLevel">Lowest heading level to include. Defaults to 3.</param>
    /// <param name="includePageNumbers">Whether to show page numbers. Defaults to <c>true</c>.</param>
    /// <param name="useHyperlinks">Whether entries link to their heading. Defaults to <c>true</c>.</param>
    /// <returns>
    /// The result, including the number of entries. An entry count of zero means the document has
    /// no paragraphs with heading styles.
    /// </returns>
    [ServiceAction("insert-toc")]
    FieldResult InsertTableOfContents(
        IWordBatch batch,
        int? paragraphIndex = null,
        int upperHeadingLevel = 1,
        int lowerHeadingLevel = 3,
        bool includePageNumbers = true,
        bool useHyperlinks = true);

    /// <summary>
    /// Updates every table of contents in the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>The number of tables of contents updated.</returns>
    [ServiceAction("update-toc")]
    FieldResult UpdateTableOfContents(IWordBatch batch);

    /// <summary>
    /// Updates all fields in the main text, headers and footers.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>The number of fields updated.</returns>
    [ServiceAction("update-all")]
    FieldResult UpdateAll(IWordBatch batch);

    /// <summary>
    /// Inserts a page number field into the header or footer of every section.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="position">Either <c>footer</c> (default) or <c>header</c>.</param>
    /// <param name="alignment">One of <c>left</c>, <c>center</c> (default) or <c>right</c>.</param>
    /// <param name="includeTotalPages">
    /// Whether to render "Page X of Y" instead of just the page number.
    /// </param>
    /// <returns>The number of page number fields inserted.</returns>
    [ServiceAction("insert-page-number")]
    FieldResult InsertPageNumber(
        IWordBatch batch,
        string position = "footer",
        string alignment = "center",
        bool includeTotalPages = false);
}

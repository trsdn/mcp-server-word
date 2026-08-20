using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.List;

/// <summary>
/// List operations: bullets, numbering, levels and restarting numbering.
/// </summary>
[ServiceCategory("list", "List")]
[McpTool("list",
    Title = "List Operations",
    Description = "Bullet and numbered list operations on an open document. "
        + "list(apply, session_id, start_index=2, end_index=5, list_type='number') turns a range of "
        + "paragraphs into a numbered list; list_type can be bullet, number or outline-number. "
        + "list(set-level, session_id, start_index=3, end_index=4, level=2) indents paragraphs to a "
        + "sub-level; levels run from 1 to 9 and only outline-number lists render a distinct format "
        + "per level. "
        + "list(restart, session_id, start_index=6) starts the numbering over at that paragraph, "
        + "which is how two separate numbered lists are kept apart. "
        + "list(remove, session_id, start_index=2, end_index=5) strips the list formatting again. "
        + "list(get, session_id) reports the list formatting of every paragraph, including the "
        + "bullet or number Word renders. "
        + "Paragraph indexes are 1-based and shift whenever paragraphs are added or removed, so "
        + "read them with paragraph(list) right before using them. "
        + "Omitting end_index applies the action to start_index alone.")]
public interface IListCommands
{
    /// <summary>
    /// Reports the list formatting of the paragraphs of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="listedOnly">
    /// Whether to return only paragraphs that carry list formatting. Defaults to true.
    /// </param>
    /// <returns>The matching paragraphs.</returns>
    [ServiceAction("get")]
    ListResult Get(IWordBatch batch, bool listedOnly = true);

    /// <summary>
    /// Formats a range of paragraphs as a bullet or numbered list.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="startIndex">1-based index of the first paragraph.</param>
    /// <param name="endIndex">
    /// 1-based index of the last paragraph. Defaults to the start paragraph.
    /// </param>
    /// <param name="listType">
    /// The kind of list: bullet, number or outline-number. Defaults to bullet.
    /// </param>
    /// <param name="level">The list level to apply, from 1 to 9. Defaults to 1.</param>
    /// <param name="continuePreviousList">
    /// Whether a numbered list continues the numbering of the preceding list instead of starting at
    /// 1. Defaults to false, because continuing an unrelated earlier list is rarely intended.
    /// </param>
    /// <returns>The formatted paragraphs.</returns>
    [ServiceAction("apply")]
    ListResult Apply(
        IWordBatch batch,
        int startIndex,
        int? endIndex = null,
        string listType = "bullet",
        int level = 1,
        bool continuePreviousList = false);

    /// <summary>
    /// Sets the list level of a range of paragraphs.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="startIndex">1-based index of the first paragraph.</param>
    /// <param name="level">The list level, from 1 to 9.</param>
    /// <param name="endIndex">
    /// 1-based index of the last paragraph. Defaults to the start paragraph.
    /// </param>
    /// <returns>The changed paragraphs.</returns>
    [ServiceAction("set-level")]
    ListResult SetLevel(IWordBatch batch, int startIndex, int level, int? endIndex = null);

    /// <summary>
    /// Restarts the numbering of a list at a paragraph.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="startIndex">1-based index of the paragraph the numbering restarts at.</param>
    /// <param name="endIndex">
    /// 1-based index of the last paragraph of the restarted list. Defaults to the end of the list
    /// the start paragraph belongs to.
    /// </param>
    /// <returns>The changed paragraphs.</returns>
    [ServiceAction("restart")]
    ListResult Restart(IWordBatch batch, int startIndex, int? endIndex = null);

    /// <summary>
    /// Removes the list formatting of a range of paragraphs.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="startIndex">1-based index of the first paragraph.</param>
    /// <param name="endIndex">
    /// 1-based index of the last paragraph. Defaults to the start paragraph.
    /// </param>
    /// <returns>The changed paragraphs.</returns>
    [ServiceAction("remove")]
    ListResult Remove(IWordBatch batch, int startIndex, int? endIndex = null);
}

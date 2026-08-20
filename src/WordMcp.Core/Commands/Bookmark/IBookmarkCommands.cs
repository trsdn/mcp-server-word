using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Bookmark;

/// <summary>
/// Bookmark operations: listing, adding, reading and deleting named positions in a document.
/// </summary>
[ServiceCategory("bookmark", "Bookmark")]
[McpTool("bookmark",
    Title = "Bookmark Operations",
    Description = "Bookmark operations on an open document. "
        + "bookmark(list, session_id) returns every bookmark with its name, paragraph index and a "
        + "preview of the bookmarked text. "
        + "bookmark(add, session_id, name=\"Intro\", paragraph_index=2) bookmarks a paragraph; pass "
        + "end_paragraph_index to span several paragraphs, or anchor_text to bookmark just a phrase "
        + "inside the paragraph. "
        + "bookmark(get-text, session_id, name=\"Intro\") returns the full bookmarked text, which is "
        + "the reliable way to re-read a passage after edits have shifted every index. "
        + "bookmark(delete, session_id, name=\"Intro\") removes the bookmark; the text stays. "
        + "Bookmark names must start with a letter and may only contain letters, digits and "
        + "underscores - Word rejects spaces and punctuation with an unhelpful error, so they are "
        + "checked before the call reaches Word. "
        + "Bookmarks survive edits elsewhere in the document, which makes them the stable way to "
        + "refer to a passage across several calls.")]
public interface IBookmarkCommands
{
    /// <summary>
    /// Lists the bookmarks of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="maxTextLength">Maximum preview length per bookmark.</param>
    /// <returns>The bookmarks.</returns>
    [ServiceAction("list")]
    BookmarkListResult List(IWordBatch batch, int maxTextLength = 200);

    /// <summary>
    /// Adds a bookmark to a paragraph, a paragraph range or a phrase inside a paragraph.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="name">The bookmark name.</param>
    /// <param name="paragraphIndex">1-based index of the first paragraph.</param>
    /// <param name="endParagraphIndex">1-based index of the last paragraph. Single paragraph when omitted.</param>
    /// <param name="anchorText">Bookmark only this phrase inside the paragraph. Whole paragraph when omitted.</param>
    /// <returns>The bookmark that was added.</returns>
    [ServiceAction("add")]
    BookmarkResult Add(
        IWordBatch batch,
        string name,
        int paragraphIndex,
        int? endParagraphIndex = null,
        string? anchorText = null);

    /// <summary>
    /// Reads the text of a bookmark.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="name">The bookmark name.</param>
    /// <returns>The bookmarked text.</returns>
    [ServiceAction("get-text")]
    BookmarkTextResult GetText(IWordBatch batch, string name);

    /// <summary>
    /// Deletes a bookmark. The bookmarked text is kept.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="name">The bookmark name.</param>
    /// <returns>The number of bookmarks left.</returns>
    [ServiceAction("delete")]
    BookmarkResult Delete(IWordBatch batch, string name);
}

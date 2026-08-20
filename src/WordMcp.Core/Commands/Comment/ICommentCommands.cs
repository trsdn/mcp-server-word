using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Comment;

/// <summary>
/// Comment operations: listing, adding, resolving and deleting comments.
/// </summary>
[ServiceCategory("comment", "Comment")]
[McpTool("comment",
    Title = "Comment Operations",
    Description = "Comment operations on an open document, the basis for review workflows. "
        + "comment(list, session_id) returns every comment with its author, date, text and the "
        + "document text it refers to. "
        + "comment(add, session_id, paragraph_index=3, text='Please shorten this') attaches a "
        + "comment to a whole paragraph; pass anchor_text to attach it to a phrase inside that "
        + "paragraph instead. "
        + "comment(resolve, session_id, index=1) marks a comment as done; pass resolved=false to "
        + "reopen it. "
        + "comment(delete, session_id, index=1) removes a comment. "
        + "Comment indexes are 1-based and shift after every delete, so when removing several "
        + "comments work from the highest index downwards. "
        + "Paragraph indexes come from paragraph(list).")]
public interface ICommentCommands
{
    /// <summary>
    /// Lists the comments of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="unresolvedOnly">
    /// Whether to return only comments that are not marked as resolved. Defaults to false.
    /// </param>
    /// <returns>The matching comments.</returns>
    [ServiceAction("list")]
    CommentListResult List(IWordBatch batch, bool unresolvedOnly = false);

    /// <summary>
    /// Adds a comment to a paragraph or to a phrase inside it.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="paragraphIndex">1-based index of the paragraph to comment on.</param>
    /// <param name="text">The comment text.</param>
    /// <param name="anchorText">
    /// A phrase inside the paragraph the comment should attach to. The whole paragraph when omitted.
    /// </param>
    /// <returns>The created comment.</returns>
    [ServiceAction("add")]
    CommentResult Add(IWordBatch batch, int paragraphIndex, string text, string? anchorText = null);

    /// <summary>
    /// Marks a comment as resolved or reopens it.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the comment.</param>
    /// <param name="resolved">Whether the comment is resolved. Defaults to true.</param>
    /// <returns>The changed comment.</returns>
    [ServiceAction("resolve")]
    CommentResult Resolve(IWordBatch batch, int index, bool resolved = true);

    /// <summary>
    /// Deletes a comment.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the comment.</param>
    /// <returns>The number of comments left.</returns>
    [ServiceAction("delete")]
    CommentResult Delete(IWordBatch batch, int index);
}

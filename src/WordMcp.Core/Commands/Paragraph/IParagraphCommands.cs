using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Paragraph;

/// <summary>
/// Paragraph-level operations: listing, adding, deleting, styling and alignment.
/// </summary>
[ServiceCategory("paragraph", "Paragraph")]
[McpTool("paragraph",
    Title = "Paragraph Operations",
    Description = "List paragraphs with their styles, add or insert paragraphs, delete them, and set style or alignment.")]
public interface IParagraphCommands
{
    /// <summary>
    /// Lists paragraphs of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="skip">Number of paragraphs to skip.</param>
    /// <param name="take">Maximum number of paragraphs to return.</param>
    /// <param name="includeEmpty">Whether empty paragraphs are included.</param>
    /// <returns>The requested paragraphs.</returns>
    [ServiceAction("list")]
    ParagraphListResult List(IWordBatch batch, int skip = 0, int take = 200, bool includeEmpty = true);

    /// <summary>
    /// Appends a paragraph at the end of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="text">The paragraph text.</param>
    /// <param name="style">Optional style name, for example <c>Heading 1</c>.</param>
    /// <param name="alignment">Optional alignment: left, center, right or justify.</param>
    /// <returns>The created paragraph.</returns>
    [ServiceAction("add")]
    ParagraphResult Add(IWordBatch batch, string text, string? style = null, string? alignment = null);

    /// <summary>
    /// Inserts a paragraph before an existing paragraph.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the paragraph to insert before.</param>
    /// <param name="text">The paragraph text.</param>
    /// <param name="style">Optional style name.</param>
    /// <param name="alignment">Optional alignment: left, center, right or justify.</param>
    /// <returns>The created paragraph.</returns>
    [ServiceAction("insert")]
    ParagraphResult Insert(IWordBatch batch, int index, string text, string? style = null, string? alignment = null);

    /// <summary>
    /// Deletes a paragraph.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the paragraph to delete.</param>
    /// <returns>The result of the delete operation.</returns>
    [ServiceAction("delete")]
    ParagraphResult Delete(IWordBatch batch, int index);

    /// <summary>
    /// Applies a style to a paragraph.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based paragraph index.</param>
    /// <param name="style">Style name, for example <c>Heading 2</c> or <c>Normal</c>.</param>
    /// <returns>The updated paragraph.</returns>
    [ServiceAction("set-style")]
    ParagraphResult SetStyle(IWordBatch batch, int index, string style);

    /// <summary>
    /// Sets the alignment of a paragraph.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based paragraph index.</param>
    /// <param name="alignment">Alignment: left, center, right or justify.</param>
    /// <returns>The updated paragraph.</returns>
    [ServiceAction("set-alignment")]
    ParagraphResult SetAlignment(IWordBatch batch, int index, string alignment);
}

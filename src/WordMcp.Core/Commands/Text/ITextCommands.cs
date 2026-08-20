using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Text;

/// <summary>
/// Text-level operations: reading, appending, searching, replacing and formatting.
/// </summary>
[ServiceCategory("text", "Text")]
[McpTool("text",
    Title = "Text Operations",
    Description = "Read document text, append text, find and replace, and apply character formatting.")]
public interface ITextCommands
{
    /// <summary>
    /// Reads the text of the document or of a character range.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="start">Optional 0-based start character position.</param>
    /// <param name="end">Optional exclusive end character position.</param>
    /// <param name="maxLength">Maximum number of characters to return.</param>
    /// <returns>The requested text.</returns>
    [ServiceAction("get")]
    TextResult Get(IWordBatch batch, int? start = null, int? end = null, int maxLength = 100_000);

    /// <summary>
    /// Appends text at the end of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="text">The text to append.</param>
    /// <param name="newParagraph">Whether to start a new paragraph before appending.</param>
    /// <returns>The appended text and its position.</returns>
    [ServiceAction("append")]
    TextResult Append(IWordBatch batch, string text, bool newParagraph = true);

    /// <summary>
    /// Finds occurrences of a search term.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="searchText">The term to search for.</param>
    /// <param name="matchCase">Whether the search is case sensitive.</param>
    /// <param name="matchWholeWord">Whether only whole words match.</param>
    /// <param name="maxResults">Maximum number of matches to return.</param>
    /// <returns>The matches found.</returns>
    [ServiceAction("find")]
    FindResult Find(
        IWordBatch batch,
        string searchText,
        bool matchCase = false,
        bool matchWholeWord = false,
        int maxResults = 100);

    /// <summary>
    /// Replaces occurrences of a search term.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="searchText">The term to search for.</param>
    /// <param name="replaceText">The replacement text.</param>
    /// <param name="matchCase">Whether the search is case sensitive.</param>
    /// <param name="matchWholeWord">Whether only whole words match.</param>
    /// <param name="replaceAll">Whether to replace all occurrences or only the first.</param>
    /// <returns>The number of replacements performed.</returns>
    [ServiceAction("replace")]
    ReplaceResult Replace(
        IWordBatch batch,
        string searchText,
        string replaceText,
        bool matchCase = false,
        bool matchWholeWord = false,
        bool replaceAll = true);

    /// <summary>
    /// Applies character formatting to a character range.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="start">0-based start character position.</param>
    /// <param name="end">Exclusive end character position.</param>
    /// <param name="bold">Whether to set or clear bold.</param>
    /// <param name="italic">Whether to set or clear italic.</param>
    /// <param name="underline">Whether to set or clear underline.</param>
    /// <param name="fontName">Font family name to apply.</param>
    /// <param name="fontSize">Font size in points to apply.</param>
    /// <param name="color">Font colour in hex notation, for example <c>#0078D4</c>.</param>
    /// <returns>The result of the formatting operation.</returns>
    [ServiceAction("format")]
    OperationResult Format(
        IWordBatch batch,
        int start,
        int end,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? fontName = null,
        double? fontSize = null,
        string? color = null);
}

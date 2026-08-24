using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Footnote;

/// <summary>
/// Footnote and endnote operations.
/// </summary>
/// <remarks>
/// Footnotes live in their own story, which is why <c>text(get)</c> never returns them: a document
/// with footnotes looks complete to a reader of the body text while part of its content is missing.
/// Every action takes a <c>kind</c> of <c>footnote</c> or <c>endnote</c>, because the two are
/// separate collections in Word with independent numbering.
/// </remarks>
[ServiceCategory("footnote", "Footnote")]
[McpTool("footnote",
    Title = "Footnote and Endnote Operations",
    Description = "Footnote and endnote operations on an open document. "
        + "IMPORTANT: text(get) does NOT include footnotes - they live in a separate story - so a "
        + "document that carries its sources in footnotes looks complete while it is not. Call "
        + "footnote(list, session_id) before summarising or translating a document. "
        + "footnote(add, session_id, paragraph_index=3, text=\"See annual report 2025, p. 14.\") "
        + "appends a reference mark at the end of that paragraph; pass anchor_text to attach it "
        + "directly after a phrase instead. "
        + "footnote(set-text, session_id, index=2, text=\"...\") rewrites one note. "
        + "footnote(delete, session_id, index=2) removes the note and its mark, and Word renumbers "
        + "the rest - so when deleting several, work from the highest index downwards. "
        + "kind='endnote' addresses endnotes, a separate collection with its own numbering; every "
        + "action defaults to footnotes.")]
public interface IFootnoteCommands
{
    /// <summary>
    /// Lists the footnotes or endnotes of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="kind">Either <c>footnote</c> or <c>endnote</c>.</param>
    /// <param name="maxTextLength">Maximum note text length per entry.</param>
    /// <returns>The notes.</returns>
    [ServiceAction("list")]
    FootnoteListResult List(IWordBatch batch, string kind = "footnote", int maxTextLength = 500);

    /// <summary>
    /// Adds a footnote or endnote to a paragraph.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="paragraphIndex">1-based index of the paragraph to attach the mark to.</param>
    /// <param name="text">The note text.</param>
    /// <param name="kind">Either <c>footnote</c> or <c>endnote</c>.</param>
    /// <param name="anchorText">
    /// Place the mark directly after this phrase inside the paragraph. Appended at the end of the
    /// paragraph when omitted.
    /// </param>
    /// <returns>The note that was added.</returns>
    [ServiceAction("add")]
    FootnoteResult Add(
        IWordBatch batch,
        int paragraphIndex,
        string text,
        string kind = "footnote",
        string? anchorText = null);

    /// <summary>
    /// Replaces the text of a footnote or endnote.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the note.</param>
    /// <param name="text">The new note text.</param>
    /// <param name="kind">Either <c>footnote</c> or <c>endnote</c>.</param>
    /// <returns>The note after the change.</returns>
    [ServiceAction("set-text")]
    FootnoteResult SetText(IWordBatch batch, int index, string text, string kind = "footnote");

    /// <summary>
    /// Deletes a footnote or endnote together with its reference mark.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based index of the note.</param>
    /// <param name="kind">Either <c>footnote</c> or <c>endnote</c>.</param>
    /// <returns>The number of notes left.</returns>
    [ServiceAction("delete")]
    FootnoteResult Delete(IWordBatch batch, int index, string kind = "footnote");
}

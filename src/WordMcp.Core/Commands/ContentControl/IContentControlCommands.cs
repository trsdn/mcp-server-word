using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.ContentControl;

/// <summary>
/// Content control operations, the structured fields Word templates are built from.
/// </summary>
/// <remarks>
/// A content control is addressed by its <c>tag</c> or by its Word <c>id</c>. Tags are not unique,
/// and a template that repeats a field - a customer name in the header and again in the signature
/// block - relies on that. So <c>set-text</c> by tag updates every match and reports how many, while
/// <c>id</c> always addresses exactly one control.
/// </remarks>
[ServiceCategory("contentControl", "ContentControl")]
[McpTool("content-control",
    Title = "Content Control Operations",
    Description = "Read and fill the content controls of a Word template. Content controls are the "
        + "structured fields a template exposes - they carry a tag, a title and a type - and filling "
        + "them is the reliable way to complete a template, because find-and-replace on the body "
        + "text corrupts the control and loses the binding. "
        + "content-control(list, session_id) shows every control with its tag, type and current "
        + "text; start here, the tags are what you address afterwards. "
        + "content-control(set-text, session_id, tag=\"CustomerName\", text=\"Contoso Ltd\") fills "
        + "EVERY control with that tag, which is what a template that repeats a field expects; pass "
        + "id=... instead to hit exactly one. "
        + "content-control(get, session_id, tag=\"CustomerName\") reads them back. "
        + "content-control(add, session_id, paragraph_index=2, type=\"rich-text\", tag=\"Notes\") "
        + "creates a new control around a paragraph. "
        + "content-control(delete, session_id, tag=\"Draft\", delete_contents=false) removes the "
        + "control and by default keeps its text in the document. "
        + "A checkbox is set with is_checked=true rather than text, and a control whose contents are "
        + "locked is reported as such instead of being silently skipped.")]
public interface IContentControlCommands
{
    /// <summary>
    /// Lists the content controls of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="tag">Restrict the listing to controls carrying this tag.</param>
    /// <param name="maxTextLength">Maximum text length per entry.</param>
    /// <returns>The content controls.</returns>
    [ServiceAction("list")]
    ContentControlListResult List(IWordBatch batch, string? tag = null, int maxTextLength = 500);

    /// <summary>
    /// Reads the content controls addressed by tag or id.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="tag">The tag to look for.</param>
    /// <param name="id">The Word identifier of a single control.</param>
    /// <param name="maxTextLength">Maximum text length per entry.</param>
    /// <returns>The matching content controls.</returns>
    [ServiceAction("get")]
    ContentControlResult Get(IWordBatch batch, string? tag = null, string? id = null, int maxTextLength = 5000);

    /// <summary>
    /// Fills the content controls addressed by tag or id.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="text">The text to write. Ignored for checkbox controls.</param>
    /// <param name="tag">The tag to look for; every match is updated.</param>
    /// <param name="id">The Word identifier of a single control.</param>
    /// <param name="isChecked">Tick state for checkbox controls.</param>
    /// <returns>The controls after the change.</returns>
    [ServiceAction("set-text")]
    ContentControlResult SetText(
        IWordBatch batch,
        string? text = null,
        string? tag = null,
        string? id = null,
        bool? isChecked = null);

    /// <summary>
    /// Wraps a paragraph in a new content control.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="paragraphIndex">1-based index of the paragraph to wrap.</param>
    /// <param name="type">The control type, for example <c>rich-text</c> or <c>text</c>.</param>
    /// <param name="tag">The tag to assign.</param>
    /// <param name="title">The title shown in Word.</param>
    /// <returns>The control that was created.</returns>
    [ServiceAction("add")]
    ContentControlResult Add(
        IWordBatch batch,
        int paragraphIndex,
        string type = "rich-text",
        string? tag = null,
        string? title = null);

    /// <summary>
    /// Removes the content controls addressed by tag or id.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="tag">The tag to look for; every match is removed.</param>
    /// <param name="id">The Word identifier of a single control.</param>
    /// <param name="deleteContents">Also delete the text inside the control.</param>
    /// <returns>The controls that were removed.</returns>
    [ServiceAction("delete")]
    ContentControlResult Delete(IWordBatch batch, string? tag = null, string? id = null, bool deleteContents = false);
}

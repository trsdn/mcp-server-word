using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Style;

/// <summary>
/// Style operations: listing, creating, modifying and deleting styles.
/// </summary>
[ServiceCategory("style", "Style")]
[McpTool("style",
    Title = "Style Operations",
    Description = "Style operations on an open document. Styles keep formatting consistent "
        + "without touching every piece of text. "
        + "style(list, session_id) returns only the styles the document actually uses; pass "
        + "in_use_only=false for the complete list, which is over 370 entries on a localized Word. "
        + "style(create, session_id, name='Callout', style_type='paragraph', base_style='Normal') "
        + "adds a custom style. "
        + "style(modify, session_id, name='Callout', font_name='Calibri', font_size=11, bold=true, "
        + "color='#C00000', alignment='center', space_after=12) changes formatting; omitted "
        + "properties stay as they are. "
        + "style(delete, session_id, name='Callout') removes a custom style; built-in styles "
        + "cannot be deleted. "
        + "LOCALIZED NAMES: Word reports styles under localized names, so a German Word calls "
        + "'Heading 1' 'Ueberschrift 1'. Built-in styles are addressable by their English name; "
        + "list returns both, and english_name is the one to send back. "
        + "Sizes and spacing are in points (72 pt = 1 inch).")]
public interface IStyleCommands
{
    /// <summary>
    /// Lists the styles of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="inUseOnly">
    /// Whether to return only styles applied or modified in the document. Defaults to true,
    /// because a localized Word defines several hundred styles that are never used.
    /// </param>
    /// <param name="styleType">
    /// Restrict the result to one kind: paragraph, character, table or list. All kinds when omitted.
    /// </param>
    /// <returns>The matching styles.</returns>
    [ServiceAction("list")]
    StyleListResult List(IWordBatch batch, bool inUseOnly = true, string? styleType = null);

    /// <summary>
    /// Creates a custom style.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="name">Name of the new style; it must not exist yet.</param>
    /// <param name="styleType">Kind of style: paragraph, character, table or list. Defaults to paragraph.</param>
    /// <param name="baseStyle">Style to inherit from, for example <c>Normal</c>.</param>
    /// <returns>The created style.</returns>
    [ServiceAction("create")]
    StyleResult Create(IWordBatch batch, string name, string styleType = "paragraph", string? baseStyle = null);

    /// <summary>
    /// Changes the formatting of a style. Only the properties supplied are written.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="name">Name of the style to change; English names work for built-in styles.</param>
    /// <param name="fontName">New font name.</param>
    /// <param name="fontSize">New font size in points.</param>
    /// <param name="bold">Whether the style is bold.</param>
    /// <param name="italic">Whether the style is italic.</param>
    /// <param name="underline">Whether the style is underlined.</param>
    /// <param name="color">Font colour as a hex value such as <c>#C00000</c>.</param>
    /// <param name="alignment">Paragraph alignment: left, center, right or justify.</param>
    /// <param name="spaceBefore">Space above a paragraph in points.</param>
    /// <param name="spaceAfter">Space below a paragraph in points.</param>
    /// <param name="lineSpacing">Line spacing in points.</param>
    /// <returns>The style after the change.</returns>
    [ServiceAction("modify")]
    StyleResult Modify(
        IWordBatch batch,
        string name,
        string? fontName = null,
        double? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? color = null,
        string? alignment = null,
        double? spaceBefore = null,
        double? spaceAfter = null,
        double? lineSpacing = null);

    /// <summary>
    /// Deletes a custom style. Built-in styles are rejected, because Word cannot remove them.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="name">Name of the style to delete.</param>
    /// <returns>The result of the operation.</returns>
    [ServiceAction("delete")]
    StyleResult Delete(IWordBatch batch, string name);
}

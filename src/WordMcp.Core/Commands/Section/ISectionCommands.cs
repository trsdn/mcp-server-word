using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Section;

/// <summary>
/// Section operations: listing sections, adding section breaks and changing the page setup.
/// </summary>
/// <remarks>
/// A section owns the page setup and the headers and footers. A document always has at least one
/// section, so <c>list</c> never returns an empty result. All measurements are in points
/// (72 pt = 1 inch), which is the unit Word uses internally.
/// </remarks>
[ServiceCategory("section", "Section")]
[McpTool("section",
    Title = "Section Operations",
    Description = "List sections, insert section breaks and change margins, orientation and paper size.")]
public interface ISectionCommands
{
    /// <summary>
    /// Lists all sections with their page setup.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>The sections of the document.</returns>
    [ServiceAction("list")]
    SectionListResult List(IWordBatch batch);

    /// <summary>
    /// Inserts a section break, which creates a new section.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="startType">
    /// How the new section starts: <c>next-page</c>, <c>continuous</c>, <c>even-page</c> or
    /// <c>odd-page</c>. Defaults to <c>next-page</c>.
    /// </param>
    /// <param name="paragraphIndex">
    /// 1-based paragraph to insert the break after. When omitted the break goes to the end of the
    /// document.
    /// </param>
    /// <returns>The new section.</returns>
    [ServiceAction("add")]
    SectionResult Add(IWordBatch batch, string startType = "next-page", int? paragraphIndex = null);

    /// <summary>
    /// Changes margins, orientation and paper size.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="sectionIndex">
    /// 1-based section to change. When omitted the change applies to the whole document.
    /// </param>
    /// <param name="topMargin">Top margin in points.</param>
    /// <param name="bottomMargin">Bottom margin in points.</param>
    /// <param name="leftMargin">Left margin in points.</param>
    /// <param name="rightMargin">Right margin in points.</param>
    /// <param name="orientation">Either <c>portrait</c> or <c>landscape</c>.</param>
    /// <param name="paperSize">A name such as <c>a4</c>, <c>letter</c> or <c>legal</c>.</param>
    /// <returns>The page setup after the change.</returns>
    [ServiceAction("page-setup")]
    SectionResult PageSetup(
        IWordBatch batch,
        int? sectionIndex = null,
        double? topMargin = null,
        double? bottomMargin = null,
        double? leftMargin = null,
        double? rightMargin = null,
        string? orientation = null,
        string? paperSize = null);
}

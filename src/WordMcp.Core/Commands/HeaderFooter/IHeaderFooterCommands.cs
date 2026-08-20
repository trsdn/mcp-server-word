using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.HeaderFooter;

/// <summary>
/// Header and footer operations, per section and per page type.
/// </summary>
/// <remarks>
/// Headers and footers belong to a section, not to the document. Two Word behaviours shape this
/// API: a new section inherits its headers from the previous one until the link is broken, and the
/// <c>first-page</c> and <c>even-pages</c> variants are only rendered once the matching switch is
/// enabled on the section. Writing takes care of both.
/// </remarks>
[ServiceCategory("headerFooter", "HeaderFooter")]
[McpTool("header-footer",
    Title = "Header and Footer Operations",
    Description = "Read, write and clear headers and footers per section and page type.")]
public interface IHeaderFooterCommands
{
    /// <summary>
    /// Reads headers or footers.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="kind">Either <c>header</c> or <c>footer</c>.</param>
    /// <param name="sectionIndex">1-based section. When omitted all sections are returned.</param>
    /// <param name="type">
    /// <c>primary</c>, <c>first-page</c> or <c>even-pages</c>. Defaults to <c>primary</c>.
    /// </param>
    /// <returns>The matching headers or footers.</returns>
    [ServiceAction("get")]
    HeaderFooterListResult Get(
        IWordBatch batch,
        string kind = "header",
        int? sectionIndex = null,
        string type = "primary");

    /// <summary>
    /// Writes a header or footer.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="text">The text to write.</param>
    /// <param name="kind">Either <c>header</c> or <c>footer</c>.</param>
    /// <param name="sectionIndex">1-based section. When omitted every section is written.</param>
    /// <param name="type">
    /// <c>primary</c>, <c>first-page</c> or <c>even-pages</c>. Defaults to <c>primary</c>. The
    /// matching section switch is enabled automatically, because Word would otherwise store the
    /// text without ever rendering it.
    /// </param>
    /// <param name="alignment">Optional <c>left</c>, <c>center</c>, <c>right</c> or <c>justify</c>.</param>
    /// <returns>The headers or footers after the change.</returns>
    [ServiceAction("set")]
    HeaderFooterResult Set(
        IWordBatch batch,
        string text,
        string kind = "header",
        int? sectionIndex = null,
        string type = "primary",
        string? alignment = null);

    /// <summary>
    /// Clears the content of a header or footer.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="kind">Either <c>header</c> or <c>footer</c>.</param>
    /// <param name="sectionIndex">1-based section. When omitted every section is cleared.</param>
    /// <param name="type"><c>primary</c>, <c>first-page</c> or <c>even-pages</c>.</param>
    /// <returns>The headers or footers after the change.</returns>
    [ServiceAction("clear")]
    HeaderFooterResult Clear(
        IWordBatch batch,
        string kind = "header",
        int? sectionIndex = null,
        string type = "primary");
}

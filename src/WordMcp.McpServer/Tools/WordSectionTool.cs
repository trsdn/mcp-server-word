using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>section</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordSectionAction>))]
public enum WordSectionAction
{
    /// <summary>List all sections.</summary>
    [JsonStringEnumMemberName("list")] List,

    /// <summary>Insert a section break.</summary>
    [JsonStringEnumMemberName("add")] Add,

    /// <summary>Change margins, orientation and paper size.</summary>
    [JsonStringEnumMemberName("page-setup")] PageSetup
}

/// <summary>
/// Section operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordSectionTool
{
    /// <summary>
    /// Lists sections, inserts section breaks and changes the page setup.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="start_type">How a new section starts.</param>
    /// <param name="paragraph_index">1-based paragraph to insert the break after.</param>
    /// <param name="section_index">1-based section for page-setup; all sections when omitted.</param>
    /// <param name="top_margin">Top margin in points.</param>
    /// <param name="bottom_margin">Bottom margin in points.</param>
    /// <param name="left_margin">Left margin in points.</param>
    /// <param name="right_margin">Right margin in points.</param>
    /// <param name="orientation">Either <c>portrait</c> or <c>landscape</c>.</param>
    /// <param name="paper_size">A name such as <c>a4</c>, <c>letter</c> or <c>legal</c>.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "section", Title = "Section Operations")]
    [Description("Section operations on an open document. A section owns the page setup and the "
        + "headers and footers, so changing margins or orientation for part of a document means "
        + "adding a section first. "
        + "section(list, session_id) returns index, start type, margins, page size and orientation. "
        + "section(add, session_id, start_type='next-page', paragraph_index=...) inserts a section "
        + "break after that paragraph, or at the end when omitted. "
        + "section(page-setup, session_id, section_index=2, orientation='landscape', top_margin=72) "
        + "changes one section; without section_index it changes the whole document. "
        + "ALL MEASUREMENTS ARE IN POINTS: 72 pt = 1 inch = 2.54 cm.")]
    public static string Section(
        WordSectionAction action,
        string session_id,
        [DefaultValue("next-page")] string start_type = "next-page",
        [DefaultValue(null)] int? paragraph_index = null,
        [DefaultValue(null)] int? section_index = null,
        [DefaultValue(null)] double? top_margin = null,
        [DefaultValue(null)] double? bottom_margin = null,
        [DefaultValue(null)] double? left_margin = null,
        [DefaultValue(null)] double? right_margin = null,
        [DefaultValue(null)] string? orientation = null,
        [DefaultValue(null)] string? paper_size = null)
        => WordToolsBase.Execute("section", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordSectionAction.List => WordServices.Sections.List(batch),
                WordSectionAction.Add => WordServices.Sections.Add(batch, start_type, paragraph_index),
                WordSectionAction.PageSetup => WordServices.Sections.PageSetup(
                    batch,
                    section_index,
                    top_margin,
                    bottom_margin,
                    left_margin,
                    right_margin,
                    orientation,
                    paper_size),
                _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
            };
        });
}

using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>table</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordTableAction>))]
public enum WordTableAction
{
    /// <summary>List all tables.</summary>
    [JsonStringEnumMemberName("list")] List,

    /// <summary>Create a table.</summary>
    [JsonStringEnumMemberName("create")] Create,

    /// <summary>Read the cells of a table.</summary>
    [JsonStringEnumMemberName("read")] Read,

    /// <summary>Write a single cell.</summary>
    [JsonStringEnumMemberName("set-cell")] SetCell,

    /// <summary>Append a row.</summary>
    [JsonStringEnumMemberName("add-row")] AddRow,

    /// <summary>Delete a row.</summary>
    [JsonStringEnumMemberName("delete-row")] DeleteRow,

    /// <summary>Apply a table style.</summary>
    [JsonStringEnumMemberName("set-style")] SetStyle
}

/// <summary>
/// Table operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordTableTool
{
    /// <summary>
    /// Lists, creates, reads and edits tables.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="index">1-based table index.</param>
    /// <param name="rows">Number of rows for create.</param>
    /// <param name="columns">Number of columns for create.</param>
    /// <param name="row">1-based row number for set-cell and delete-row.</param>
    /// <param name="column">1-based column number for set-cell.</param>
    /// <param name="text">Cell text for set-cell.</param>
    /// <param name="values">Cell values for add-row, applied left to right.</param>
    /// <param name="style">Table style name, for example <c>Table Grid</c>.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "table", Title = "Table Operations")]
    [Description("Table operations on an open document. "
        + "table(list, session_id) returns index, size and style of every table. "
        + "table(create, session_id, rows=3, columns=4, style='Table Grid') appends a table. "
        + "table(read, session_id, index=1) returns all cell values as rows. "
        + "table(set-cell, session_id, index=1, row=1, column=2, text='Total') writes one cell. "
        + "table(add-row, session_id, index=1, values=['a','b']) appends a filled row. "
        + "table(delete-row|set-style, session_id, index=1, ...) edits an existing table. "
        + "All indexes are 1-based.")]
    public static string Table(
        WordTableAction action,
        string session_id,
        [DefaultValue(null)] int? index = null,
        [DefaultValue(null)] int? rows = null,
        [DefaultValue(null)] int? columns = null,
        [DefaultValue(null)] int? row = null,
        [DefaultValue(null)] int? column = null,
        [DefaultValue(null)] string? text = null,
        [DefaultValue(null)] string[]? values = null,
        [DefaultValue(null)] string? style = null)
        => WordToolsBase.Execute("table", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordTableAction.List => WordServices.Tables.List(batch),
                WordTableAction.Create => WordServices.Tables.Create(
                    batch, Require(rows, nameof(rows)), Require(columns, nameof(columns)), style),
                WordTableAction.Read => WordServices.Tables.Read(batch, Require(index, nameof(index))),
                WordTableAction.SetCell => WordServices.Tables.SetCell(
                    batch,
                    Require(index, nameof(index)),
                    Require(row, nameof(row)),
                    Require(column, nameof(column)),
                    text ?? string.Empty),
                WordTableAction.AddRow => WordServices.Tables.AddRow(
                    batch, Require(index, nameof(index)), values),
                WordTableAction.DeleteRow => WordServices.Tables.DeleteRow(
                    batch, Require(index, nameof(index)), Require(row, nameof(row))),
                WordTableAction.SetStyle => WordServices.Tables.SetStyle(
                    batch, Require(index, nameof(index)), RequireText(style, nameof(style))),
                _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
            };
        });

    private static int Require(int? value, string name)
        => value ?? throw new ArgumentException($"{name} is required for this action.", name);

    private static string RequireText(string? value, string name)
        => string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException($"{name} is required for this action.", name)
            : value;
}

using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>document</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordDocumentAction>))]
public enum WordDocumentAction
{
    /// <summary>Read document statistics.</summary>
    [JsonStringEnumMemberName("get-info")] GetInfo,

    /// <summary>Read built-in document properties.</summary>
    [JsonStringEnumMemberName("get-properties")] GetProperties,

    /// <summary>Write built-in document properties.</summary>
    [JsonStringEnumMemberName("set-properties")] SetProperties,

    /// <summary>Export the document to PDF.</summary>
    [JsonStringEnumMemberName("export-pdf")] ExportPdf,

    /// <summary>Save a copy in another format.</summary>
    [JsonStringEnumMemberName("save-as")] SaveAs
}

/// <summary>
/// Document-level operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordDocumentTool
{
    /// <summary>
    /// Inspects statistics and properties of a document and exports it.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="output_path">Target path for export-pdf and save-as.</param>
    /// <param name="title">Document title for set-properties.</param>
    /// <param name="author">Author for set-properties.</param>
    /// <param name="subject">Subject for set-properties.</param>
    /// <param name="keywords">Keywords for set-properties.</param>
    /// <param name="comments">Comments for set-properties.</param>
    /// <param name="company">Company for set-properties.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "document", Title = "Document Operations")]
    [Description("Document-level operations. "
        + "document(get-info, session_id) returns word/page/paragraph/table counts. "
        + "document(get-properties|set-properties, session_id, title=..., author=...) reads or writes metadata. "
        + "document(export-pdf, session_id, output_path='C:\\\\...\\\\report.pdf') writes a PDF without changing the document. "
        + "document(save-as, session_id, output_path='C:\\\\...\\\\copy.rtf') saves a copy; the format follows the extension "
        + "(.docx, .docm, .doc, .pdf, .rtf, .txt, .html) and the open document is saved as part of the operation.")]
    public static string Document(
        WordDocumentAction action,
        string session_id,
        [DefaultValue(null)] string? output_path = null,
        [DefaultValue(null)] string? title = null,
        [DefaultValue(null)] string? author = null,
        [DefaultValue(null)] string? subject = null,
        [DefaultValue(null)] string? keywords = null,
        [DefaultValue(null)] string? comments = null,
        [DefaultValue(null)] string? company = null)
        => WordToolsBase.Execute("document", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordDocumentAction.GetInfo => WordServices.Documents.GetInfo(batch),
                WordDocumentAction.GetProperties => WordServices.Documents.GetProperties(batch),
                WordDocumentAction.SetProperties => WordServices.Documents.SetProperties(
                    batch, title, author, subject, keywords, comments, company),
                WordDocumentAction.ExportPdf => WordServices.Documents.ExportPdf(
                    batch, Require(output_path, nameof(output_path))),
                WordDocumentAction.SaveAs => WordServices.Documents.SaveAs(
                    batch, Require(output_path, nameof(output_path))),
                _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
            };
        });

    private static string Require(string? value, string name)
        => string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException($"{name} is required for this action.", name)
            : value;
}

using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>file</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordFileAction>))]
public enum WordFileAction
{
    /// <summary>Open an existing document and start a session.</summary>
    [JsonStringEnumMemberName("open")] Open,

    /// <summary>Create a new document and start a session.</summary>
    [JsonStringEnumMemberName("create")] Create,

    /// <summary>Save the document of a session.</summary>
    [JsonStringEnumMemberName("save")] Save,

    /// <summary>Close a session, optionally saving first.</summary>
    [JsonStringEnumMemberName("close")] Close,

    /// <summary>List all open sessions.</summary>
    [JsonStringEnumMemberName("list")] List,

    /// <summary>Check whether a file can be opened without starting a session.</summary>
    [JsonStringEnumMemberName("test")] Test
}

/// <summary>
/// File and session management — the entry point for every Word workflow.
/// </summary>
[McpServerToolType]
public static class WordFileTool
{
    /// <summary>
    /// Opens, creates, saves, closes and lists Word document sessions.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="path">Absolute path of the document (open, create, test).</param>
    /// <param name="session_id">Session identifier (save, close).</param>
    /// <param name="save">Whether to save when closing.</param>
    /// <param name="show">Whether the Word window is visible.</param>
    /// <param name="timeout_seconds">Per-operation timeout in seconds.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "file", Title = "File Operations", Destructive = true)]
    [Description("File and session management — the FIRST tool of every workflow. "
        + "WORKFLOW: file(open, path='C:\\\\...\\\\report.docx') -> use session_id with document/text/paragraph/table tools -> file(close, save=true). "
        + "NEW FILES: file(create, path='C:\\\\...\\\\new.docx') creates the file AND starts a session. "
        + "REUSE: call file(list) first — if the document is already open, reuse its session_id. "
        + "The file must be CLOSED in the Word desktop app; COM requires exclusive access. "
        + "show=true makes Word visible. timeout_seconds: max time per operation (default 300).")]
    public static string File(
        WordFileAction action,
        [DefaultValue(null)] string? path = null,
        [DefaultValue(null)] string? session_id = null,
        [DefaultValue(false)] bool save = false,
        [DefaultValue(false)] bool show = false,
        [DefaultValue(300)] int timeout_seconds = 300)
        => WordToolsBase.Execute("file", action.ToString().ToLowerInvariant(), () => action switch
        {
            WordFileAction.Open => Open(path, show, timeout_seconds),
            WordFileAction.Create => Create(path, show, timeout_seconds),
            WordFileAction.Save => ServiceBridge.Invoke("session.save", session_id),
            WordFileAction.Close => ServiceBridge.Invoke("session.close", session_id, new { save }),
            WordFileAction.List => ServiceBridge.Invoke("session.list"),
            WordFileAction.Test => Test(path),
            _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
        });

    private static object Open(string? path, bool show, int timeoutSeconds)
        => ServiceBridge.Invoke("session.open", args: new
        {
            filePath = WordToolsBase.ValidatePath(path, mustExist: true),
            visible = show,
            timeoutSeconds
        });

    private static object Create(string? path, bool show, int timeoutSeconds)
        => ServiceBridge.Invoke("session.create", args: new
        {
            filePath = WordToolsBase.ValidatePath(path, mustExist: false),
            visible = show,
            timeoutSeconds
        });

    private static object Test(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("path is required for 'test'.", nameof(path));
        }

        // Unlike open and create, an unsupported extension is reported as a finding rather than
        // thrown: answering "what would happen" is the entire point of this action.
        return ServiceBridge.Invoke("session.test", args: new { filePath = Path.GetFullPath(path) });
    }
}

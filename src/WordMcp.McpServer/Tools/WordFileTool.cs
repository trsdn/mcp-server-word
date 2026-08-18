using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;
using WordMcp.ComInterop;

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
            WordFileAction.Save => Save(session_id),
            WordFileAction.Close => Close(session_id, save),
            WordFileAction.List => List(),
            WordFileAction.Test => Test(path),
            _ => throw new ArgumentException($"Unknown action '{action}'.", nameof(action))
        });

    private static object Open(string? path, bool show, int timeoutSeconds)
    {
        var fullPath = WordToolsBase.ValidatePath(path, mustExist: true);

        var existing = WordServices.Sessions.FindByPath(fullPath);
        if (existing != null)
        {
            return new
            {
                success = true,
                sessionId = existing.SessionId,
                filePath = existing.FilePath,
                reused = true,
                message = "Document is already open; reusing the existing session."
            };
        }

        var info = WordServices.Sessions.Open(fullPath, show, timeoutSeconds);
        return new
        {
            success = true,
            sessionId = info.SessionId,
            filePath = info.FilePath,
            visible = info.Visible,
            reused = false,
            message = "Session opened. Pass session_id to the other tools."
        };
    }

    private static object Create(string? path, bool show, int timeoutSeconds)
    {
        var fullPath = WordToolsBase.ValidatePath(path, mustExist: false);

        if (System.IO.File.Exists(fullPath))
        {
            throw new IOException(
                $"File already exists: {fullPath}. Use file(action:'open') instead, or choose a different path.");
        }

        var info = WordServices.Sessions.Create(fullPath, show, timeoutSeconds);
        return new
        {
            success = true,
            sessionId = info.SessionId,
            filePath = info.FilePath,
            visible = info.Visible,
            message = "Document created and session opened."
        };
    }

    private static object Save(string? sessionId)
    {
        WordToolsBase.Batch(sessionId);
        WordServices.Sessions.Save(sessionId!);
        return new { success = true, sessionId, message = "Document saved." };
    }

    private static object Close(string? sessionId, bool save)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ArgumentException("session_id is required for 'close'.", nameof(sessionId));
        }

        var closed = WordServices.Sessions.Close(sessionId, save);
        return new
        {
            success = closed,
            sessionId,
            saved = save && closed,
            message = closed
                ? save ? "Document saved and session closed." : "Session closed without saving."
                : $"No open session with id '{sessionId}'."
        };
    }

    private static object List()
    {
        var sessions = WordServices.Sessions.List()
            .Select(s => new
            {
                sessionId = s.SessionId,
                filePath = s.FilePath,
                visible = s.Visible,
                openedAt = s.OpenedAt
            })
            .ToArray();

        return new { success = true, count = sessions.Length, sessions };
    }

    private static object Test(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("path is required for 'test'.", nameof(path));
        }

        var fullPath = Path.GetFullPath(path);
        var extension = Path.GetExtension(fullPath);
        var exists = System.IO.File.Exists(fullPath);

        var problems = new List<string>();

        if (!exists)
        {
            problems.Add("File does not exist.");
        }

        if (!ComInteropConstants.SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            problems.Add($"Unsupported extension '{extension}'.");
        }

        if (exists)
        {
            try
            {
                FileAccessValidator.ValidateFileNotLocked(fullPath);
            }
            catch (InvalidOperationException ex)
            {
                problems.Add(ex.Message);
            }

            if (FileAccessValidator.IsIrmProtected(fullPath))
            {
                problems.Add("File appears to be IRM/RMS protected and cannot be automated.");
            }
        }

        return new
        {
            success = problems.Count == 0,
            filePath = fullPath,
            exists,
            extension,
            canOpen = problems.Count == 0,
            problems,
            message = problems.Count == 0 ? "File can be opened." : string.Join(" ", problems)
        };
    }
}

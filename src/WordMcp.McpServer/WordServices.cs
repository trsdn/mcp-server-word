using Microsoft.Extensions.Logging;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Document;
using WordMcp.Core.Commands.Bookmark;
using WordMcp.Core.Commands.Comment;
using WordMcp.Core.Commands.Screenshot;
using WordMcp.Core.Commands.Field;
using WordMcp.Core.Commands.HeaderFooter;
using WordMcp.Core.Commands.Image;
using WordMcp.Core.Commands.List;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Revision;
using WordMcp.Core.Commands.Section;
using WordMcp.Core.Commands.Style;
using WordMcp.Core.Commands.Table;
using WordMcp.Core.Commands.Text;

namespace WordMcp.McpServer;

/// <summary>
/// Process-wide singletons used by the MCP tools.
/// </summary>
/// <remarks>
/// The MCP SDK discovers tools as static methods, so the session manager and command
/// implementations are held here instead of being resolved through DI.
/// </remarks>
internal static class WordServices
{
    private static readonly Lock Gate = new();
    private static SessionManager? _sessions;

    /// <summary>Gets the shared session manager.</summary>
    public static SessionManager Sessions
    {
        get
        {
            lock (Gate)
            {
                return _sessions ??= new SessionManager(Logger);
            }
        }
    }

    /// <summary>Gets or sets the logger handed to new sessions.</summary>
    public static ILogger? Logger { get; set; }

    /// <summary>Gets the document command implementation.</summary>
    public static IDocumentCommands Documents { get; } = new DocumentCommands();

    /// <summary>Gets the text command implementation.</summary>
    public static ITextCommands Texts { get; } = new TextCommands();

    /// <summary>Gets the paragraph command implementation.</summary>
    public static IParagraphCommands Paragraphs { get; } = new ParagraphCommands();

    /// <summary>Gets the table command implementation.</summary>
    public static ITableCommands Tables { get; } = new TableCommands();

    /// <summary>Gets the image command implementation.</summary>
    public static IImageCommands Images { get; } = new ImageCommands();

    /// <summary>Gets the field command implementation.</summary>
    public static IFieldCommands Fields { get; } = new FieldCommands();

    /// <summary>Gets the section command implementation.</summary>
    public static ISectionCommands Sections { get; } = new SectionCommands();

    /// <summary>Gets the header and footer command implementation.</summary>
    public static IHeaderFooterCommands HeadersFooters { get; } = new HeaderFooterCommands();

    /// <summary>Gets the style command implementation.</summary>
    public static IStyleCommands Styles { get; } = new StyleCommands();

    /// <summary>Gets the list command implementation.</summary>
    public static IListCommands Lists { get; } = new ListCommands();

    /// <summary>Gets the comment command implementation.</summary>
    public static ICommentCommands Comments { get; } = new CommentCommands();

    /// <summary>Gets the revision command implementation.</summary>
    public static IRevisionCommands Revisions { get; } = new RevisionCommands();

    /// <summary>Gets the bookmark command implementation.</summary>
    public static IBookmarkCommands Bookmarks { get; } = new BookmarkCommands();

    /// <summary>Gets the screenshot command implementation.</summary>
    public static IScreenshotCommands Screenshots { get; } = new ScreenshotCommands();

    /// <summary>
    /// Saves and closes every open session. Called on shutdown so unsaved work is not lost.
    /// </summary>
    public static void Shutdown()
    {
        lock (Gate)
        {
            _sessions?.Dispose();
            _sessions = null;
        }
    }
}

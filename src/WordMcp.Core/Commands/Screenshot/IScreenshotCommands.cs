using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Screenshot;

/// <summary>
/// Page rendering: turning a page of the open document into a PNG.
/// </summary>
[ServiceCategory("screenshot", "Screenshot")]
[McpTool("screenshot",
    Title = "Screenshot Operations",
    Description = "Renders a page of an open document as a PNG image. "
        + "screenshot(page, session_id) renders page 1; pass page=3 for another page. "
        + "Use this to check the layout visually - page breaks, table widths, image placement and "
        + "header positions are far easier to judge from the rendered page than from measurements. "
        + "The image is written to a file and the path is returned. Pass include_image=true to also "
        + "get the PNG inline as base64, which is only worth it when the image is actually going to "
        + "be looked at, since it is large. "
        + "dpi defaults to 150, which is readable without being wasteful; 96 is enough for a rough "
        + "layout check and 300 approaches print quality. "
        + "Rendering goes through a PDF export of the single page, so unsaved changes are included.")]
public interface IScreenshotCommands
{
    /// <summary>
    /// Renders one page of the document as a PNG.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="page">1-based page number.</param>
    /// <param name="outputPath">Where to write the PNG. A temporary file when omitted.</param>
    /// <param name="dpi">Rendering resolution between 36 and 600.</param>
    /// <param name="includeImage">Whether to also return the PNG inline as base64.</param>
    /// <returns>The rendered page.</returns>
    [ServiceAction("page")]
    ScreenshotResult Page(
        IWordBatch batch,
        int page = 1,
        string? outputPath = null,
        int dpi = 150,
        bool includeImage = false);
}

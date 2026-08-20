using WordMcp.ComInterop.Session;
using WordMcp.Core.Attributes;
using WordMcp.Core.Models;

namespace WordMcp.Core.Commands.Image;

/// <summary>
/// Image operations on the inline images of a document.
/// </summary>
/// <remarks>
/// Word keeps images in two separate collections: <c>InlineShapes</c> sit in the text flow like a
/// character, <c>Shapes</c> float freely with their own anchor. These commands cover
/// <c>InlineShapes</c> only, which is what inserting a picture produces by default.
/// </remarks>
[ServiceCategory("image", "Image")]
[McpTool("image",
    Title = "Image Operations",
    Description = "Inline image operations on an open document. "
        + "image(list, session_id) returns index, size and alt text of every image. "
        + "image(insert, session_id, image_path='C:/pics/chart.png', width=300, caption='Figure 1') "
        + "appends an image; pass paragraph_index to place it before a specific paragraph. "
        + "image(resize, session_id, index=1, scale_percent=50) or (index=1, width=200) resizes. "
        + "image(replace, session_id, index=1, image_path=...) swaps the picture but keeps position and size. "
        + "image(delete|set-alt-text, session_id, index=1, ...) edits an existing image. "
        + "Sizes are in points (1 point = 1/72 inch), not pixels. All indexes are 1-based. "
        + "Only inline images are covered, not floating shapes.")]
public interface IImageCommands
{
    /// <summary>
    /// Lists all inline images of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <returns>The images of the document.</returns>
    [ServiceAction("list")]
    ImageListResult List(IWordBatch batch);

    /// <summary>
    /// Inserts an image, by default at the end of the document.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="imagePath">Absolute path of the image file.</param>
    /// <param name="paragraphIndex">
    /// 1-based paragraph to insert before. When omitted the image is appended.
    /// </param>
    /// <param name="width">Optional width in points.</param>
    /// <param name="height">Optional height in points.</param>
    /// <param name="caption">Optional caption placed in a paragraph below the image.</param>
    /// <param name="altText">Optional alternative text for screen readers.</param>
    /// <returns>The inserted image.</returns>
    [ServiceAction("insert")]
    ImageResult Insert(
        IWordBatch batch,
        string imagePath,
        int? paragraphIndex = null,
        double? width = null,
        double? height = null,
        string? caption = null,
        string? altText = null);

    /// <summary>
    /// Resizes an image, either to an absolute size in points or by a percentage.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based image index.</param>
    /// <param name="width">Target width in points.</param>
    /// <param name="height">Target height in points.</param>
    /// <param name="scalePercent">Target size as a percentage of the current size.</param>
    /// <param name="lockAspectRatio">
    /// Whether Word adjusts the other dimension automatically. Defaults to <c>true</c>.
    /// </param>
    /// <returns>The resized image.</returns>
    [ServiceAction("resize")]
    ImageResult Resize(
        IWordBatch batch,
        int index,
        double? width = null,
        double? height = null,
        double? scalePercent = null,
        bool lockAspectRatio = true);

    /// <summary>
    /// Replaces the picture of an image while keeping its position and size.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based image index.</param>
    /// <param name="imagePath">Absolute path of the new image file.</param>
    /// <param name="keepSize">Whether to restore the previous size. Defaults to <c>true</c>.</param>
    /// <returns>The replaced image.</returns>
    [ServiceAction("replace")]
    ImageResult Replace(IWordBatch batch, int index, string imagePath, bool keepSize = true);

    /// <summary>
    /// Deletes an image.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based image index.</param>
    /// <returns>The result of the operation.</returns>
    [ServiceAction("delete")]
    ImageResult Delete(IWordBatch batch, int index);

    /// <summary>
    /// Sets the alternative text of an image.
    /// </summary>
    /// <param name="batch">The Word batch to operate on.</param>
    /// <param name="index">1-based image index.</param>
    /// <param name="altText">The alternative text.</param>
    /// <returns>The updated image.</returns>
    [ServiceAction("set-alt-text")]
    ImageResult SetAltText(IWordBatch batch, int index, string altText);
}

using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace WordMcp.McpServer.Tools;

/// <summary>Actions of the <c>image</c> tool.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<WordImageAction>))]
public enum WordImageAction
{
    /// <summary>List all inline images.</summary>
    [JsonStringEnumMemberName("list")] List,

    /// <summary>Insert an image.</summary>
    [JsonStringEnumMemberName("insert")] Insert,

    /// <summary>Resize an image.</summary>
    [JsonStringEnumMemberName("resize")] Resize,

    /// <summary>Replace the picture of an image.</summary>
    [JsonStringEnumMemberName("replace")] Replace,

    /// <summary>Delete an image.</summary>
    [JsonStringEnumMemberName("delete")] Delete,

    /// <summary>Set the alternative text of an image.</summary>
    [JsonStringEnumMemberName("set-alt-text")] SetAltText
}

/// <summary>
/// Image operations for an open Word session.
/// </summary>
[McpServerToolType]
public static class WordImageTool
{
    /// <summary>
    /// Inserts, lists and edits inline images.
    /// </summary>
    /// <param name="action">The action to perform.</param>
    /// <param name="session_id">Session identifier from file(open) or file(create).</param>
    /// <param name="image_path">Absolute path of the image file for insert and replace.</param>
    /// <param name="index">1-based image index.</param>
    /// <param name="paragraph_index">1-based paragraph to insert before; appends when omitted.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    /// <param name="scale_percent">Size as a percentage of the current size, for resize.</param>
    /// <param name="lock_aspect_ratio">Whether resizing keeps the aspect ratio.</param>
    /// <param name="keep_size">Whether replace restores the previous size.</param>
    /// <param name="caption">Caption placed below the image on insert.</param>
    /// <param name="alt_text">Alternative text for screen readers.</param>
    /// <returns>A JSON payload describing the result.</returns>
    [McpServerTool(Name = "image", Title = "Image Operations")]
    [Description("Inline image operations on an open document. "
        + "image(list, session_id) returns index, size and alt text of every image. "
        + "image(insert, session_id, image_path='C:/pics/chart.png', width=300, caption='Figure 1') "
        + "appends an image; pass paragraph_index to place it before a specific paragraph. "
        + "image(resize, session_id, index=1, scale_percent=50) or (index=1, width=200) resizes. "
        + "image(replace, session_id, index=1, image_path=...) swaps the picture but keeps position and size. "
        + "image(delete|set-alt-text, session_id, index=1, ...) edits an existing image. "
        + "Sizes are in points (1 point = 1/72 inch), not pixels. All indexes are 1-based. "
        + "Only inline images are covered, not floating shapes.")]
    public static string Image(
        WordImageAction action,
        string session_id,
        [DefaultValue(null)] string? image_path = null,
        [DefaultValue(null)] int? index = null,
        [DefaultValue(null)] int? paragraph_index = null,
        [DefaultValue(null)] double? width = null,
        [DefaultValue(null)] double? height = null,
        [DefaultValue(null)] double? scale_percent = null,
        [DefaultValue(true)] bool lock_aspect_ratio = true,
        [DefaultValue(true)] bool keep_size = true,
        [DefaultValue(null)] string? caption = null,
        [DefaultValue(null)] string? alt_text = null)
        => WordToolsBase.Execute("image", action.ToString(), () =>
        {
            var batch = WordToolsBase.Batch(session_id);

            return action switch
            {
                WordImageAction.List => WordServices.Images.List(batch),
                WordImageAction.Insert => WordServices.Images.Insert(
                    batch,
                    RequireText(image_path, nameof(image_path)),
                    paragraph_index,
                    width,
                    height,
                    caption,
                    alt_text),
                WordImageAction.Resize => WordServices.Images.Resize(
                    batch,
                    Require(index, nameof(index)),
                    width,
                    height,
                    scale_percent,
                    lock_aspect_ratio),
                WordImageAction.Replace => WordServices.Images.Replace(
                    batch,
                    Require(index, nameof(index)),
                    RequireText(image_path, nameof(image_path)),
                    keep_size),
                WordImageAction.Delete => WordServices.Images.Delete(
                    batch, Require(index, nameof(index))),
                WordImageAction.SetAltText => WordServices.Images.SetAltText(
                    batch, Require(index, nameof(index)), alt_text ?? string.Empty),
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

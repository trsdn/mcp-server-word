using WordMcp.ComInterop;

namespace WordMcp.Core.Utilities;

/// <summary>
/// Conversions between the friendly string values used by MCP tools and Word COM enumerations.
/// </summary>
public static class WordConversions
{
    /// <summary>
    /// Converts an alignment name to the matching <c>WdParagraphAlignment</c> value.
    /// </summary>
    /// <param name="alignment">One of <c>left</c>, <c>center</c>, <c>right</c>, <c>justify</c>.</param>
    /// <returns>The Word alignment constant.</returns>
    /// <exception cref="ArgumentException">The alignment name is unknown.</exception>
    public static int ToWdAlignment(string alignment)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(alignment);

        return alignment.Trim().ToLowerInvariant() switch
        {
            "left" => ComInteropConstants.WdAlignParagraphLeft,
            "center" or "centre" or "centered" => ComInteropConstants.WdAlignParagraphCenter,
            "right" => ComInteropConstants.WdAlignParagraphRight,
            "justify" or "justified" => ComInteropConstants.WdAlignParagraphJustify,
            _ => throw new ArgumentException(
                $"Unknown alignment '{alignment}'. Use one of: left, center, right, justify.", nameof(alignment))
        };
    }

    /// <summary>
    /// Converts a <c>WdParagraphAlignment</c> value to its friendly name.
    /// </summary>
    /// <param name="wdAlignment">The Word alignment constant.</param>
    /// <returns>The friendly alignment name.</returns>
    public static string FromWdAlignment(int wdAlignment) => wdAlignment switch
    {
        ComInteropConstants.WdAlignParagraphLeft => "left",
        ComInteropConstants.WdAlignParagraphCenter => "center",
        ComInteropConstants.WdAlignParagraphRight => "right",
        ComInteropConstants.WdAlignParagraphJustify => "justify",
        _ => "other"
    };

    /// <summary>
    /// Maps a file extension to the matching <c>WdSaveFormat</c> value.
    /// </summary>
    /// <param name="path">Target file path; only its extension is inspected.</param>
    /// <returns>The Word save format constant.</returns>
    /// <exception cref="ArgumentException">The extension has no supported save format.</exception>
    public static int ToWdSaveFormat(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".docx" => ComInteropConstants.WdFormatXmlDocument,
            ".docm" => ComInteropConstants.WdFormatXmlDocumentMacroEnabled,
            ".doc" => ComInteropConstants.WdFormatDocument97,
            ".pdf" => ComInteropConstants.WdFormatPdf,
            ".rtf" => ComInteropConstants.WdFormatRtf,
            ".txt" => ComInteropConstants.WdFormatText,
            ".html" or ".htm" => ComInteropConstants.WdFormatFilteredHtml,
            var ext => throw new ArgumentException(
                $"Unsupported target format '{ext}'. Use .docx, .docm, .doc, .pdf, .rtf, .txt or .html.", nameof(path))
        };
    }

    /// <summary>
    /// Removes the control characters Word appends to range text.
    /// </summary>
    /// <param name="rangeText">Raw text read from a Word range.</param>
    /// <returns>The cleaned text.</returns>
    /// <remarks>
    /// Paragraph ranges end with a carriage return and table cell ranges end with the
    /// end-of-cell marker. Both are artefacts of the object model, not content.
    /// </remarks>
    public static string CleanRangeText(string? rangeText)
        => rangeText?.TrimEnd('\r', '\a', '\u0007') ?? string.Empty;

    /// <summary>
    /// Parses a hex colour such as <c>#0078D4</c> into the BGR integer Word expects.
    /// </summary>
    /// <param name="hexColor">Colour in <c>#RRGGBB</c> or <c>RRGGBB</c> notation.</param>
    /// <returns>The Word colour value.</returns>
    /// <exception cref="ArgumentException">The value is not a six-digit hex colour.</exception>
    public static int ToWdColor(string hexColor)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hexColor);

        var value = hexColor.TrimStart('#');
        if (value.Length != 6 || !int.TryParse(value, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out int rgb))
        {
            throw new ArgumentException(
                $"Invalid colour '{hexColor}'. Use hex notation such as #0078D4.", nameof(hexColor));
        }

        int r = (rgb >> 16) & 0xFF;
        int g = (rgb >> 8) & 0xFF;
        int b = rgb & 0xFF;

        // Word stores colours as BGR, not RGB.
        return (b << 16) | (g << 8) | r;
    }
}

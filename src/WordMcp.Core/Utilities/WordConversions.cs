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

    /// <summary>
    /// Converts an orientation name to the matching <c>WdOrientation</c> value.
    /// </summary>
    /// <param name="orientation">Either <c>portrait</c> or <c>landscape</c>.</param>
    /// <returns>The Word orientation constant.</returns>
    /// <exception cref="ArgumentException">The orientation name is unknown.</exception>
    public static int ToWdOrientation(string orientation)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(orientation);

        return orientation.Trim().ToLowerInvariant() switch
        {
            "portrait" or "vertical" => ComInteropConstants.WdOrientPortrait,
            "landscape" or "horizontal" => ComInteropConstants.WdOrientLandscape,
            _ => throw new ArgumentException(
                $"Unknown orientation '{orientation}'. Use portrait or landscape.", nameof(orientation))
        };
    }

    /// <summary>
    /// Converts a <c>WdOrientation</c> value to its friendly name.
    /// </summary>
    /// <param name="wdOrientation">The Word orientation constant.</param>
    /// <returns>The friendly orientation name.</returns>
    public static string FromWdOrientation(int wdOrientation) => wdOrientation switch
    {
        ComInteropConstants.WdOrientPortrait => "portrait",
        ComInteropConstants.WdOrientLandscape => "landscape",
        _ => "other"
    };

    /// <summary>
    /// Converts a paper size name to the matching <c>WdPaperSize</c> value.
    /// </summary>
    /// <param name="paperSize">A name such as <c>a4</c>, <c>letter</c> or <c>legal</c>.</param>
    /// <returns>The Word paper size constant.</returns>
    /// <exception cref="ArgumentException">The paper size name is unknown.</exception>
    public static int ToWdPaperSize(string paperSize)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(paperSize);

        return paperSize.Trim().ToLowerInvariant().Replace(" ", string.Empty, StringComparison.Ordinal) switch
        {
            "a3" => ComInteropConstants.WdPaperA3,
            "a4" => ComInteropConstants.WdPaperA4,
            "a5" => ComInteropConstants.WdPaperA5,
            "letter" => ComInteropConstants.WdPaperLetter,
            "legal" => ComInteropConstants.WdPaperLegal,
            "tabloid" => ComInteropConstants.WdPaperTabloid,
            _ => throw new ArgumentException(
                $"Unknown paper size '{paperSize}'. Use one of: a3, a4, a5, letter, legal, tabloid.",
                nameof(paperSize))
        };
    }

    /// <summary>
    /// Converts a section start name to the matching <c>WdSectionStart</c> value.
    /// </summary>
    /// <param name="startType">
    /// One of <c>next-page</c>, <c>continuous</c>, <c>even-page</c>, <c>odd-page</c> or
    /// <c>new-column</c>.
    /// </param>
    /// <returns>The Word section start constant.</returns>
    /// <exception cref="ArgumentException">The section start name is unknown.</exception>
    public static int ToWdSectionStart(string startType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(startType);

        return NormalizeName(startType) switch
        {
            "next-page" or "nextpage" or "page" => ComInteropConstants.WdSectionNewPage,
            "continuous" => ComInteropConstants.WdSectionContinuous,
            "even-page" or "evenpage" => ComInteropConstants.WdSectionEvenPage,
            "odd-page" or "oddpage" => ComInteropConstants.WdSectionOddPage,
            "new-column" or "newcolumn" or "column" => ComInteropConstants.WdSectionNewColumn,
            _ => throw new ArgumentException(
                $"Unknown section start '{startType}'. Use one of: next-page, continuous, even-page, "
                + "odd-page, new-column.",
                nameof(startType))
        };
    }

    /// <summary>
    /// Converts a <c>WdSectionStart</c> value to its friendly name.
    /// </summary>
    /// <param name="wdSectionStart">The Word section start constant.</param>
    /// <returns>The friendly section start name.</returns>
    public static string FromWdSectionStart(int wdSectionStart) => wdSectionStart switch
    {
        ComInteropConstants.WdSectionContinuous => "continuous",
        ComInteropConstants.WdSectionNewColumn => "new-column",
        ComInteropConstants.WdSectionNewPage => "next-page",
        ComInteropConstants.WdSectionEvenPage => "even-page",
        ComInteropConstants.WdSectionOddPage => "odd-page",
        _ => "other"
    };

    /// <summary>
    /// Converts a section start name to the <c>WdBreakType</c> used when inserting the break.
    /// </summary>
    /// <param name="startType">Same values as <see cref="ToWdSectionStart"/>, except new-column.</param>
    /// <returns>The Word break type constant.</returns>
    /// <exception cref="ArgumentException">The section start name is unknown here.</exception>
    public static int ToWdSectionBreak(string startType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(startType);

        return NormalizeName(startType) switch
        {
            "next-page" or "nextpage" or "page" => ComInteropConstants.WdSectionBreakNextPage,
            "continuous" => ComInteropConstants.WdSectionBreakContinuous,
            "even-page" or "evenpage" => ComInteropConstants.WdSectionBreakEvenPage,
            "odd-page" or "oddpage" => ComInteropConstants.WdSectionBreakOddPage,
            _ => throw new ArgumentException(
                $"Unknown section break '{startType}'. Use one of: next-page, continuous, even-page, odd-page.",
                nameof(startType))
        };
    }

    /// <summary>
    /// Converts a header or footer type name to the matching <c>WdHeaderFooterIndex</c> value.
    /// </summary>
    /// <param name="type">One of <c>primary</c>, <c>first-page</c> or <c>even-pages</c>.</param>
    /// <returns>The Word header/footer index constant.</returns>
    /// <exception cref="ArgumentException">The type name is unknown.</exception>
    public static int ToWdHeaderFooterIndex(string type)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(type);

        return NormalizeName(type) switch
        {
            "primary" or "default" or "main" => ComInteropConstants.WdHeaderFooterPrimary,
            "first-page" or "firstpage" or "first" => ComInteropConstants.WdHeaderFooterFirstPage,
            "even-pages" or "evenpages" or "even" => ComInteropConstants.WdHeaderFooterEvenPages,
            _ => throw new ArgumentException(
                $"Unknown header/footer type '{type}'. Use one of: primary, first-page, even-pages.",
                nameof(type))
        };
    }

    /// <summary>
    /// Converts a <c>WdHeaderFooterIndex</c> value to its friendly name.
    /// </summary>
    /// <param name="wdHeaderFooterIndex">The Word header/footer index constant.</param>
    /// <returns>The friendly type name.</returns>
    public static string FromWdHeaderFooterIndex(int wdHeaderFooterIndex) => wdHeaderFooterIndex switch
    {
        ComInteropConstants.WdHeaderFooterPrimary => "primary",
        ComInteropConstants.WdHeaderFooterFirstPage => "first-page",
        ComInteropConstants.WdHeaderFooterEvenPages => "even-pages",
        _ => "other"
    };

    /// <summary>
    /// Converts a style kind name to the matching <c>WdStyleType</c> value.
    /// </summary>
    /// <param name="styleType">One of <c>paragraph</c>, <c>character</c>, <c>table</c> or <c>list</c>.</param>
    /// <returns>The Word style type constant.</returns>
    /// <exception cref="ArgumentException">The style kind is unknown.</exception>
    public static int ToWdStyleType(string styleType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleType);

        return NormalizeName(styleType) switch
        {
            "paragraph" => ComInteropConstants.WdStyleTypeParagraph,
            "character" => ComInteropConstants.WdStyleTypeCharacter,
            "table" => ComInteropConstants.WdStyleTypeTable,
            "list" => ComInteropConstants.WdStyleTypeList,
            _ => throw new ArgumentException(
                $"Unknown style type '{styleType}'. Use one of: paragraph, character, table, list.",
                nameof(styleType))
        };
    }

    /// <summary>
    /// Converts a <c>WdStyleType</c> value to its friendly name.
    /// </summary>
    /// <param name="wdStyleType">The Word style type constant.</param>
    /// <returns>The friendly style kind name.</returns>
    public static string FromWdStyleType(int wdStyleType) => wdStyleType switch
    {
        ComInteropConstants.WdStyleTypeParagraph => "paragraph",
        ComInteropConstants.WdStyleTypeCharacter => "character",
        ComInteropConstants.WdStyleTypeTable => "table",
        ComInteropConstants.WdStyleTypeList => "list",
        _ => "other"
    };

    private static string NormalizeName(string value)
        => value.Trim().ToLowerInvariant().Replace('_', '-').Replace(" ", string.Empty, StringComparison.Ordinal);
}

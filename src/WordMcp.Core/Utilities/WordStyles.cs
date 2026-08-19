namespace WordMcp.Core.Utilities;

/// <summary>
/// Resolves English built-in style names to Word's language-independent style identifiers.
/// </summary>
/// <remarks>
/// Word stores built-in styles under localized names: on a German installation "Heading 1" is
/// called "Überschrift 1" and does not resolve by its English name. Clients (and language models)
/// however address styles in English, so built-in names are translated to their
/// <c>WdBuiltinStyle</c> constant, which every Word language accepts. Unknown names are passed
/// through unchanged so custom and localized styles keep working.
/// </remarks>
public static class WordStyles
{
    private static readonly Dictionary<string, int> BuiltInStyles =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Normal"] = -1,
            ["Heading 1"] = -2,
            ["Heading 2"] = -3,
            ["Heading 3"] = -4,
            ["Heading 4"] = -5,
            ["Heading 5"] = -6,
            ["Heading 6"] = -7,
            ["Heading 7"] = -8,
            ["Heading 8"] = -9,
            ["Heading 9"] = -10,
            ["Header"] = -32,
            ["Footer"] = -33,
            ["Caption"] = -35,
            ["Title"] = -63,
            ["Body Text"] = -67,
            ["Subtitle"] = -75,
            ["Hyperlink"] = -86,
            ["Strong"] = -88,
            ["Emphasis"] = -89,
            ["Table Grid"] = -155,
            ["List Paragraph"] = -180,
            ["Quote"] = -181,
            ["Intense Quote"] = -182,
        };

    /// <summary>
    /// Converts a style name into the value to assign to a Word <c>Style</c> property.
    /// </summary>
    /// <param name="styleName">Style name as supplied by the caller.</param>
    /// <returns>
    /// The <c>WdBuiltinStyle</c> constant when <paramref name="styleName"/> is a known English
    /// built-in name, otherwise the name itself.
    /// </returns>
    public static object Resolve(string styleName)
    {
        ArgumentNullException.ThrowIfNull(styleName);

        return BuiltInStyles.TryGetValue(styleName.Trim(), out int builtIn)
            ? builtIn
            : styleName;
    }
}

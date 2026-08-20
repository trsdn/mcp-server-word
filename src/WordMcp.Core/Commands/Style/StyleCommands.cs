using System.Runtime.InteropServices;
using WordMcp.ComInterop;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Models;
using WordMcp.Core.Utilities;

namespace WordMcp.Core.Commands.Style;

/// <summary>
/// Word COM implementation of <see cref="IStyleCommands"/>.
/// </summary>
public sealed class StyleCommands : IStyleCommands
{
    /// <inheritdoc />
    public StyleListResult List(IWordBatch batch, bool inUseOnly = true, string? styleType = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        int? wantedType = styleType is null ? null : WordConversions.ToWdStyleType(styleType);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic styles = doc.Styles;
            int total = (int)styles.Count;

            var englishNames = MapLocalNamesToEnglish(doc);
            var list = new List<StyleInfo>();

            for (int i = 1; i <= total; i++)
            {
                ct.ThrowIfCancellationRequested();

                dynamic style = styles[i];

                int type = (int)style.Type;
                if (wantedType.HasValue && type != wantedType.Value)
                    continue;

                // Word's InUse also covers built-in styles that were merely modified, which is
                // exactly the set a caller cares about when formatting a document.
                bool inUse = (bool)style.InUse;
                if (inUseOnly && !inUse)
                    continue;

                list.Add(Describe(style, englishNames));
            }

            return new StyleListResult
            {
                TotalCount = total,
                ReturnedCount = list.Count,
                Styles = list
            };
        });
    }

    /// <inheritdoc />
    public StyleResult Create(IWordBatch batch, string name, string styleType = "paragraph", string? baseStyle = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        int wdStyleType = WordConversions.ToWdStyleType(styleType);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;

            if (Find(doc, name) != null)
            {
                throw new ArgumentException(
                    $"Style '{name}' already exists. Use style(action:'modify') to change it.", nameof(name));
            }

            dynamic style = doc.Styles.Add(name, wdStyleType);

            if (!string.IsNullOrWhiteSpace(baseStyle))
            {
                try
                {
                    style.BaseStyle = WordStyles.Resolve(baseStyle);
                }
                catch (COMException ex)
                {
                    throw new ArgumentException(
                        $"Base style '{baseStyle}' does not exist in this document.", nameof(baseStyle), ex);
                }
            }

            return new StyleResult
            {
                Style = Describe(style, MapLocalNamesToEnglish(doc)),
                Message = $"Style '{name}' created."
            };
        });
    }

    /// <inheritdoc />
    public StyleResult Modify(
        IWordBatch batch,
        string name,
        string? fontName = null,
        double? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? color = null,
        string? alignment = null,
        double? spaceBefore = null,
        double? spaceAfter = null,
        double? lineSpacing = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        if (fontSize is <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(fontSize), "font_size must be greater than 0.");
        }

        if (lineSpacing is <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(lineSpacing), "line_spacing must be greater than 0.");
        }

        // Converting up front turns a bad value into a clear error before Word is involved.
        int? wdColor = color is null ? null : WordConversions.ToWdColor(color);
        int? wdAlignment = alignment is null ? null : WordConversions.ToWdAlignment(alignment);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic style = Find(doc, name)
                ?? throw new ArgumentException($"Style '{name}' does not exist in this document.", nameof(name));

            ApplyFont(style, fontName, fontSize, bold, italic, underline, wdColor);
            ApplyParagraphFormat(style, wdAlignment, spaceBefore, spaceAfter, lineSpacing);

            return new StyleResult
            {
                Style = Describe(style, MapLocalNamesToEnglish(doc)),
                Message = $"Style '{name}' modified."
            };
        });
    }

    /// <inheritdoc />
    public StyleResult Delete(IWordBatch batch, string name)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        return batch.Execute((ctx, ct) =>
        {
            dynamic doc = ctx.Document;
            dynamic style = Find(doc, name)
                ?? throw new ArgumentException($"Style '{name}' does not exist in this document.", nameof(name));

            // Word raises a generic COM error for this, so the check produces a usable message.
            if ((bool)style.BuiltIn)
            {
                throw new InvalidOperationException(
                    $"Style '{name}' is built into Word and cannot be deleted. Only custom styles can.");
            }

            style.Delete();

            return new StyleResult { Message = $"Style '{name}' deleted." };
        });
    }

    private static void ApplyFont(
        dynamic style,
        string? fontName,
        double? fontSize,
        bool? bold,
        bool? italic,
        bool? underline,
        int? wdColor)
    {
        if (fontName is null && fontSize is null && bold is null
            && italic is null && underline is null && wdColor is null)
        {
            return;
        }

        dynamic font = style.Font;

        if (!string.IsNullOrWhiteSpace(fontName))
            font.Name = fontName;

        if (fontSize.HasValue)
            font.Size = (float)fontSize.Value;

        // Word's font flags are tri-state integers, so a plain bool would not round-trip.
        if (bold.HasValue)
            font.Bold = bold.Value ? ComInteropConstants.MsoTrue : ComInteropConstants.MsoFalse;

        if (italic.HasValue)
            font.Italic = italic.Value ? ComInteropConstants.MsoTrue : ComInteropConstants.MsoFalse;

        if (underline.HasValue)
            font.Underline = underline.Value ? ComInteropConstants.MsoTrue : ComInteropConstants.MsoFalse;

        if (wdColor.HasValue)
            font.Color = wdColor.Value;
    }

    private static void ApplyParagraphFormat(
        dynamic style,
        int? wdAlignment,
        double? spaceBefore,
        double? spaceAfter,
        double? lineSpacing)
    {
        if (wdAlignment is null && spaceBefore is null && spaceAfter is null && lineSpacing is null)
        {
            return;
        }

        dynamic format;
        try
        {
            format = style.ParagraphFormat;
        }
        catch (COMException ex)
        {
            throw new ArgumentException(
                "Paragraph formatting can only be set on paragraph and table styles.", nameof(style), ex);
        }

        if (wdAlignment.HasValue)
            format.Alignment = wdAlignment.Value;

        if (spaceBefore.HasValue)
            format.SpaceBefore = (float)spaceBefore.Value;

        if (spaceAfter.HasValue)
            format.SpaceAfter = (float)spaceAfter.Value;

        if (lineSpacing.HasValue)
        {
            // LineSpacing is only honoured once the rule says the value is an exact measurement.
            format.LineSpacingRule = ComInteropConstants.WdLineSpaceExactly;
            format.LineSpacing = (float)lineSpacing.Value;
        }
    }

    /// <summary>
    /// Looks a style up by name, accepting the English name of a built-in style on a localized Word.
    /// </summary>
    private static dynamic? Find(dynamic doc, string name)
    {
        try
        {
            return doc.Styles[WordStyles.Resolve(name)];
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <summary>
    /// Builds the reverse of <see cref="WordStyles"/>: the name Word reports mapped to the English
    /// name a client can send back on any language version.
    /// </summary>
    private static Dictionary<string, string> MapLocalNamesToEnglish(dynamic doc)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in WordStyles.KnownBuiltInStyles)
        {
            try
            {
                dynamic style = doc.Styles[entry.Value];
                map[(string)style.NameLocal] = entry.Key;
            }
            catch (COMException)
            {
                // Not every built-in style exists in every template; skipping one only costs the
                // English name of that single style.
            }
        }

        return map;
    }

    private static StyleInfo Describe(dynamic style, Dictionary<string, string> englishNames)
    {
        string name = (string)style.NameLocal;
        int type = (int)style.Type;

        var info = new StyleInfo
        {
            Name = name,
            EnglishName = englishNames.TryGetValue(name, out string? english) ? english : null,
            Type = WordConversions.FromWdStyleType(type),
            BuiltIn = (bool)style.BuiltIn,
            InUse = (bool)style.InUse,
            BaseStyle = ReadBaseStyle(style)
        };

        if (type is ComInteropConstants.WdStyleTypeParagraph or ComInteropConstants.WdStyleTypeCharacter)
        {
            try
            {
                dynamic font = style.Font;
                info.FontName = (string)font.Name;
                info.FontSize = (double)(float)font.Size;
            }
            catch (COMException)
            {
                // A style can inherit its font entirely, in which case Word refuses to report one.
            }
        }

        return info;
    }

    private static string? ReadBaseStyle(dynamic style)
    {
        try
        {
            dynamic baseStyle = style.BaseStyle;
            string name = (string)baseStyle.NameLocal;
            return string.IsNullOrWhiteSpace(name) ? null : name;
        }
        catch (COMException)
        {
            // Styles without a base, such as Normal, throw instead of returning nothing.
            return null;
        }
    }
}

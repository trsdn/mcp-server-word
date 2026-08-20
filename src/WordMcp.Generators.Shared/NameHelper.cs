using System.Text;
using Microsoft.CodeAnalysis;

namespace WordMcp.Generators.Common;

/// <summary>
/// Naming conversions between the C# surface and the MCP wire format.
/// </summary>
public static class NameHelper
{
    /// <summary>Turns <c>PageSetup</c> into <c>page-setup</c>.</summary>
    public static string ToKebabCase(string pascalCase)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < pascalCase.Length; i++)
        {
            char c = pascalCase[i];
            if (char.IsUpper(c))
            {
                if (i > 0)
                {
                    sb.Append('-');
                }

                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }

    /// <summary>Turns <c>page-setup</c> into <c>PageSetup</c>.</summary>
    public static string ToPascalCase(string kebabCase)
    {
        var parts = kebabCase.Split(new[] { '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
        var sb = new StringBuilder();
        foreach (var part in parts)
        {
            sb.Append(char.ToUpperInvariant(part[0]));
            if (part.Length > 1)
            {
                sb.Append(part.Substring(1));
            }
        }

        return sb.ToString();
    }

    /// <summary>Turns <c>paragraphIndex</c> into <c>paragraph_index</c>.</summary>
    public static string ToSnakeCase(string camelCase)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < camelCase.Length; i++)
        {
            char c = camelCase[i];
            if (char.IsUpper(c))
            {
                if (i > 0)
                {
                    sb.Append('_');
                }

                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }
}

/// <summary>
/// Renders Roslyn type symbols as C# source text.
/// </summary>
public static class TypeNameHelper
{
    private static readonly SymbolDisplayFormat NullableQualifiedFormat =
        SymbolDisplayFormat.FullyQualifiedFormat.AddMiscellaneousOptions(
            SymbolDisplayMiscellaneousOptions.IncludeNullableReferenceTypeModifier);

    private static readonly HashSet<string> CollectionInterfaces = new(StringComparer.Ordinal)
    {
        "IReadOnlyList", "IReadOnlyCollection", "IEnumerable", "IList", "ICollection"
    };

    /// <summary>
    /// Renders a type as it should appear in generated source, keeping nullability intact.
    /// </summary>
    public static string GetTypeName(ITypeSymbol type, NullableAnnotation annotation = NullableAnnotation.None)
    {
        if (type.SpecialType == SpecialType.System_String)
        {
            return annotation == NullableAnnotation.Annotated ? "string?" : "string";
        }

        if (type.SpecialType == SpecialType.System_Boolean)
        {
            return "bool";
        }

        if (type.SpecialType == SpecialType.System_Int32)
        {
            return "int";
        }

        if (type.SpecialType == SpecialType.System_Double)
        {
            return "double";
        }

        if (type is INamedTypeSymbol named
            && named.IsGenericType
            && named.ConstructedFrom.SpecialType == SpecialType.System_Nullable_T)
        {
            return GetTypeName(named.TypeArguments[0]) + "?";
        }

        // Collections cross the wire as JSON arrays, and an array satisfies every read-only
        // collection interface the commands declare.
        if (type is INamedTypeSymbol collection
            && collection.IsGenericType
            && collection.TypeArguments.Length == 1
            && CollectionInterfaces.Contains(collection.ConstructedFrom.Name))
        {
            string element = GetTypeName(collection.TypeArguments[0]);
            string array = element + "[]";
            return annotation == NullableAnnotation.Annotated ? array + "?" : array;
        }

        var fullName = type.ToDisplayString(NullableQualifiedFormat);
        if (fullName.StartsWith("global::", StringComparison.Ordinal))
        {
            fullName = fullName.Substring(8);
        }

        if (annotation == NullableAnnotation.Annotated && !fullName.EndsWith("?", StringComparison.Ordinal))
        {
            fullName += "?";
        }

        return fullName;
    }

    /// <summary>
    /// Renders a parameter's default value as a C# expression.
    /// </summary>
    public static string? GetDefaultValueString(IParameterSymbol parameter)
    {
        if (!parameter.HasExplicitDefaultValue)
        {
            return null;
        }

        object? value = parameter.ExplicitDefaultValue;

        if (value is null)
        {
            return "null";
        }

        if (value is bool b)
        {
            return b ? "true" : "false";
        }

        if (value is string s)
        {
            return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        // A double literal without a decimal point becomes an int, which makes the generated
        // [DefaultValue(400)] throw InvalidCastException when the SDK reads it back as double.
        if (value is double d)
        {
            string text = d.ToString("G17", System.Globalization.CultureInfo.InvariantCulture);
            if (text.IndexOf('.') < 0 && text.IndexOf('E') < 0 && text.IndexOf('e') < 0)
            {
                text += ".0";
            }

            return text;
        }

        if (value is float f)
        {
            string text = f.ToString("G9", System.Globalization.CultureInfo.InvariantCulture);
            if (text.IndexOf('.') < 0 && text.IndexOf('E') < 0 && text.IndexOf('e') < 0)
            {
                text += ".0";
            }

            return text + "f";
        }

        return System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
    }
}

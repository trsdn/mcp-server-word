using System.Xml;
using Microsoft.CodeAnalysis;

namespace WordMcp.Generators.Common;

/// <summary>
/// Supplies the XML documentation of symbols that come from a referenced assembly.
/// </summary>
/// <remarks>
/// The command-line compiler builds metadata references without a documentation provider, so
/// <c>GetDocumentationCommentXml</c> returns nothing during a normal build even though the
/// projects generate documentation files. Reading the .xml file next to the assembly restores the
/// parameter descriptions that the generated tools hand to the language model.
/// </remarks>
public sealed class DocumentationSource
{
    private readonly Dictionary<string, string> _members = new(StringComparer.Ordinal);

    private DocumentationSource()
    {
    }

    /// <summary>An empty source, used when no documentation file could be read.</summary>
    public static DocumentationSource Empty { get; } = new();

    /// <summary>
    /// Loads the documentation files that sit next to the compilation's assembly references.
    /// </summary>
    /// <param name="compilation">The compilation being processed.</param>
    /// <param name="assemblyNamePrefix">Only references whose file name starts with this prefix are read.</param>
    /// <returns>A source covering every documentation file that could be loaded.</returns>
    public static DocumentationSource Load(Compilation compilation, string assemblyNamePrefix)
    {
        var source = new DocumentationSource();

        foreach (var reference in compilation.References)
        {
            if (reference is not PortableExecutableReference portable
                || string.IsNullOrEmpty(portable.FilePath))
            {
                continue;
            }

            string path = portable.FilePath!;
            if (!Path.GetFileName(path).StartsWith(assemblyNamePrefix, StringComparison.Ordinal))
            {
                continue;
            }

            foreach (string candidate in DocumentationPaths(path))
            {
                if (source.Add(candidate))
                {
                    break;
                }
            }
        }

        return source;
    }

    /// <summary>
    /// Yields the places the documentation of an assembly can be, most specific first. A project
    /// reference resolves to the reference assembly under <c>obj\...\ref\</c>, while the compiler
    /// writes the documentation one directory above it.
    /// </summary>
    private static IEnumerable<string> DocumentationPaths(string assemblyPath)
    {
        string fileName = Path.GetFileNameWithoutExtension(assemblyPath) + ".xml";
        string? directory = Path.GetDirectoryName(assemblyPath);

        if (string.IsNullOrEmpty(directory))
        {
            yield break;
        }

        yield return Path.Combine(directory!, fileName);

        string? parent = Path.GetDirectoryName(directory);
        if (!string.IsNullOrEmpty(parent))
        {
            yield return Path.Combine(parent!, fileName);
        }
    }

    /// <summary>
    /// Returns the documentation of a symbol, preferring what the compilation already knows.
    /// </summary>
    /// <param name="symbol">The symbol to describe.</param>
    /// <returns>The inner XML of the documentation comment, or <c>null</c>.</returns>
    public string? GetXml(ISymbol symbol)
    {
        string? xml = symbol.GetDocumentationCommentXml();
        if (!string.IsNullOrWhiteSpace(xml))
        {
            return xml;
        }

        string? id = symbol.GetDocumentationCommentId();
        if (id != null && _members.TryGetValue(id, out string? fromFile))
        {
            return fromFile;
        }

        return null;
    }

    private bool Add(string xmlPath)
    {
        try
        {
            if (!File.Exists(xmlPath))
            {
                return false;
            }

            var document = new XmlDocument();
            document.Load(xmlPath);

            var members = document.SelectNodes("/doc/members/member");
            if (members is null)
            {
                return false;
            }

            foreach (XmlNode member in members)
            {
                string? name = member.Attributes?["name"]?.Value;
                if (!string.IsNullOrEmpty(name))
                {
                    _members[name!] = member.InnerXml;
                }
            }

            return true;
        }
#pragma warning disable CA1031 // Documentation is optional, so no failure here may break the build
        catch (Exception)
        {
            // A missing or malformed documentation file only costs descriptions, not correctness.
            return false;
        }
#pragma warning restore CA1031
    }
}

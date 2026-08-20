using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.CodeAnalysis;

namespace WordMcp.Generators.Common;

/// <summary>
/// Reads a [ServiceCategory] interface and turns it into a <see cref="ServiceInfo"/>.
/// </summary>
public static class ServiceInfoExtractor
{
    /// <summary>
    /// Extracts the service description, or returns <c>null</c> when the interface is not a
    /// service category or has no MCP tool.
    /// </summary>
    /// <param name="interfaceSymbol">The candidate command interface.</param>
    /// <param name="documentation">Source of XML documentation for the interface's members.</param>
    /// <returns>The extracted service, or <c>null</c>.</returns>
    public static ServiceInfo? Extract(INamedTypeSymbol interfaceSymbol, DocumentationSource documentation)
    {
        string? category = null;
        string? pascalName = null;
        string? toolName = null;
        string? title = null;
        string? description = null;
        bool? destructive = null;

        foreach (var attribute in interfaceSymbol.GetAttributes())
        {
            switch (attribute.AttributeClass?.Name)
            {
                case "ServiceCategoryAttribute":
                    if (attribute.ConstructorArguments.Length > 0)
                    {
                        category = attribute.ConstructorArguments[0].Value?.ToString();
                    }

                    if (attribute.ConstructorArguments.Length > 1)
                    {
                        pascalName = attribute.ConstructorArguments[1].Value?.ToString();
                    }

                    break;

                case "McpToolAttribute":
                    if (attribute.ConstructorArguments.Length > 0)
                    {
                        toolName = attribute.ConstructorArguments[0].Value?.ToString();
                    }

                    foreach (var named in attribute.NamedArguments)
                    {
                        switch (named.Key)
                        {
                            case "Title":
                                title = named.Value.Value?.ToString();
                                break;
                            case "Description":
                                description = named.Value.Value?.ToString();
                                break;
                            case "Destructive":
                                if (named.Value.Value is bool flag)
                                {
                                    destructive = flag;
                                }

                                break;
                        }
                    }

                    break;
            }
        }

        if (category is null || toolName is null)
        {
            return null;
        }

        var methods = new List<MethodInfo>();

        foreach (var member in interfaceSymbol.GetMembers())
        {
            if (member is not IMethodSymbol method || method.MethodKind != MethodKind.Ordinary)
            {
                continue;
            }

            var documentationXml = ReadDocumentation(method, documentation);
            string actionName = GetActionName(method);

            var parameters = new List<ParameterInfo>();
            foreach (var parameter in method.Parameters)
            {
                // The batch is resolved from the session id, so it never reaches the wire.
                if (parameter.Type.Name == "IWordBatch")
                {
                    continue;
                }

                documentationXml.Parameters.TryGetValue(parameter.Name, out string? parameterDoc);

                parameters.Add(new ParameterInfo(
                    parameter.Name,
                    TypeNameHelper.GetTypeName(parameter.Type, parameter.NullableAnnotation),
                    parameter.HasExplicitDefaultValue,
                    TypeNameHelper.GetDefaultValueString(parameter),
                    parameterDoc));
            }

            methods.Add(new MethodInfo(
                method.Name,
                actionName,
                NameHelper.ToPascalCase(actionName),
                parameters,
                documentationXml.Summary));
        }

        if (methods.Count == 0)
        {
            return null;
        }

        return new ServiceInfo(
            category,
            pascalName ?? NameHelper.ToPascalCase(category),
            toolName,
            title,
            description,
            destructive,
            methods,
            interfaceSymbol.Name,
            interfaceSymbol.ContainingNamespace.ToDisplayString());
    }

    /// <summary>
    /// Merges the parameters of every action into the single list the generated tool method
    /// declares. When two actions share a name, the nullable variant wins because the method has
    /// to be callable for every action.
    /// </summary>
    public static List<ExposedParameter> GetExposedParameters(ServiceInfo info)
    {
        var map = new Dictionary<string, ExposedParameter>(StringComparer.Ordinal);
        var order = new List<string>();

        foreach (var method in info.Methods)
        {
            foreach (var parameter in method.Parameters)
            {
                if (!map.TryGetValue(parameter.Name, out var exposed))
                {
                    exposed = new ExposedParameter(parameter.Name, parameter.TypeName, parameter.Description)
                    {
                        DefaultValue = parameter.DefaultValue,
                        DefaultsConflict = !parameter.HasDefault
                    };
                    map[parameter.Name] = exposed;
                    order.Add(parameter.Name);
                }
                else
                {
                    if (parameter.TypeName.EndsWith("?", StringComparison.Ordinal)
                        && !exposed.TypeName.EndsWith("?", StringComparison.Ordinal))
                    {
                        exposed.TypeName = parameter.TypeName;
                    }

                    if (!parameter.HasDefault
                        || !string.Equals(parameter.DefaultValue, exposed.DefaultValue, StringComparison.Ordinal))
                    {
                        exposed.DefaultsConflict = true;
                    }

                    if (string.IsNullOrEmpty(exposed.Description) && !string.IsNullOrEmpty(parameter.Description))
                    {
                        exposed.Description = parameter.Description;
                    }
                }

                if (parameter.IsRequired)
                {
                    exposed.RequiredByActions.Add(method.ActionName);
                }
            }
        }

        var result = new List<ExposedParameter>(order.Count);
        foreach (string name in order)
        {
            var exposed = map[name];
            exposed.TotalActionCount = info.Methods.Count;
            result.Add(exposed);
        }

        return result;
    }

    private static string GetActionName(IMethodSymbol method)
    {
        foreach (var attribute in method.GetAttributes())
        {
            if (attribute.AttributeClass?.Name == "ServiceActionAttribute"
                && attribute.ConstructorArguments.Length > 0)
            {
                string? name = attribute.ConstructorArguments[0].Value?.ToString();
                if (!string.IsNullOrEmpty(name))
                {
                    return name!;
                }
            }
        }

        return NameHelper.ToKebabCase(method.Name);
    }

    /// <summary>
    /// Reads the XML documentation of a method. This only works because the projects set
    /// GenerateDocumentationFile, which puts the XML next to the assembly the generator reads.
    /// </summary>
    private static Documentation ReadDocumentation(ISymbol symbol, DocumentationSource source)
    {
        var parameters = new Dictionary<string, string>(StringComparer.Ordinal);
        string? xml = source.GetXml(symbol);

        if (string.IsNullOrEmpty(xml))
        {
            return new Documentation(null, parameters);
        }

        try
        {
            var document = new XmlDocument();
            document.LoadXml("<root>" + xml + "</root>");

            string? summary = Normalize(document.SelectSingleNode("//summary")?.InnerText);

            var nodes = document.SelectNodes("//param");
            if (nodes != null)
            {
                foreach (XmlNode node in nodes)
                {
                    string? name = node.Attributes?["name"]?.Value;
                    string? text = Normalize(node.InnerText);
                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(text))
                    {
                        parameters[name!] = text!;
                    }
                }
            }

            return new Documentation(summary, parameters);
        }
        catch (XmlException)
        {
            // Documentation is a convenience, so malformed XML must not break the build.
            return new Documentation(null, parameters);
        }
    }

    private static string? Normalize(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        return Regex.Replace(text!.Trim(), @"\s+", " ");
    }

    private sealed class Documentation
    {
        public Documentation(string? summary, Dictionary<string, string> parameters)
        {
            Summary = summary;
            Parameters = parameters;
        }

        public string? Summary { get; }

        public Dictionary<string, string> Parameters { get; }
    }
}

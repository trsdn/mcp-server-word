using System.Collections.Immutable;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;
using WordMcp.Generators.Common;

namespace WordMcp.Generators.Mcp;

/// <summary>
/// Emits one MCP tool class per command interface marked with [ServiceCategory] and [McpTool],
/// so the wire contract is derived from the interfaces instead of being restated by hand.
/// </summary>
[Generator]
public sealed class McpToolGenerator : IIncrementalGenerator
{
    private const string ToolsNamespace = "WordMcp.McpServer.Tools";
    private const string ServicesType = "WordMcp.McpServer.WordServices";

    /// <inheritdoc />
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var services = context.CompilationProvider.Select(static (compilation, _) => Discover(compilation));

        context.RegisterSourceOutput(services, static (productionContext, discovered) =>
        {
            foreach (var service in discovered)
            {
                string source = Emit(service);
                productionContext.AddSource(
                    $"WordTool.{service.Info.CategoryPascal}.g.cs",
                    SourceText.From(source, Encoding.UTF8));
            }
        });
    }

    /// <summary>
    /// Finds every command interface of the referenced Core assembly and pairs it with the
    /// WordServices property that exposes it.
    /// </summary>
    private static ImmutableArray<DiscoveredService> Discover(Compilation compilation)
    {
        var servicesByInterface = MapServiceProperties(compilation);
        var documentation = DocumentationSource.Load(compilation, "WordMcp.");
        var builder = ImmutableArray.CreateBuilder<DiscoveredService>();

        foreach (var reference in compilation.References)
        {
            if (compilation.GetAssemblyOrModuleSymbol(reference) is not IAssemblySymbol assembly)
            {
                continue;
            }

            // Only our own Core assembly can carry command interfaces; scanning the BCL would be
            // both pointless and slow.
            if (!assembly.Name.StartsWith("WordMcp.", StringComparison.Ordinal))
            {
                continue;
            }

            foreach (var type in EnumerateTypes(assembly.GlobalNamespace))
            {
                if (type.TypeKind != TypeKind.Interface)
                {
                    continue;
                }

                var info = ServiceInfoExtractor.Extract(type, documentation);
                if (info is null)
                {
                    continue;
                }

                if (!servicesByInterface.TryGetValue(info.InterfaceName, out string? property))
                {
                    property = Pluralize(info.CategoryPascal);
                }

                builder.Add(new DiscoveredService(info, property));
            }
        }

        return builder.ToImmutable();
    }

    /// <summary>
    /// Reads the static WordServices class and maps each command interface to the property that
    /// returns it, so the generated dispatch never has to guess a name.
    /// </summary>
    private static Dictionary<string, string> MapServiceProperties(Compilation compilation)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        var services = compilation.GetTypeByMetadataName(ServicesType);

        if (services is null)
        {
            return map;
        }

        foreach (var member in services.GetMembers())
        {
            if (member is IPropertySymbol { IsStatic: true } property
                && property.Type.TypeKind == TypeKind.Interface)
            {
                map[property.Type.Name] = property.Name;
            }
        }

        return map;
    }

    private static IEnumerable<INamedTypeSymbol> EnumerateTypes(INamespaceSymbol root)
    {
        foreach (var member in root.GetMembers())
        {
            if (member is INamespaceSymbol childNamespace)
            {
                foreach (var nested in EnumerateTypes(childNamespace))
                {
                    yield return nested;
                }
            }
            else if (member is INamedTypeSymbol type)
            {
                yield return type;
            }
        }
    }

    /// <summary>
    /// Fallback naming when WordServices does not expose the interface. "HeaderFooter" becomes
    /// "HeadersFooters" because each word is pluralized on its own.
    /// </summary>
    private static string Pluralize(string pascalName)
    {
        var words = SplitWords(pascalName);
        var sb = new StringBuilder();

        foreach (string word in words)
        {
            sb.Append(word);
            if (!word.EndsWith("s", StringComparison.Ordinal))
            {
                sb.Append('s');
            }
        }

        return sb.ToString();
    }

    private static List<string> SplitWords(string pascalName)
    {
        var words = new List<string>();
        var current = new StringBuilder();

        foreach (char c in pascalName)
        {
            if (char.IsUpper(c) && current.Length > 0)
            {
                words.Add(current.ToString());
                current.Clear();
            }

            current.Append(c);
        }

        if (current.Length > 0)
        {
            words.Add(current.ToString());
        }

        return words;
    }

    private static string Emit(DiscoveredService service)
    {
        var info = service.Info;
        var exposed = ServiceInfoExtractor.GetExposedParameters(info);
        string enumName = $"Word{info.CategoryPascal}Action";
        string className = $"Word{info.CategoryPascal}Tool";

        var sb = new StringBuilder();
        sb.AppendLine("// <auto-generated/>");
        sb.AppendLine("#nullable enable");
        sb.AppendLine();
        sb.AppendLine("using System;");
        sb.AppendLine("using System.ComponentModel;");
        sb.AppendLine("using System.Text.Json.Serialization;");
        sb.AppendLine("using ModelContextProtocol.Server;");
        sb.AppendLine();
        sb.AppendLine($"namespace {ToolsNamespace};");
        sb.AppendLine();

        EmitEnum(sb, info, enumName);
        EmitTool(sb, service, exposed, enumName, className);

        return sb.ToString();
    }

    private static void EmitEnum(StringBuilder sb, ServiceInfo info, string enumName)
    {
        sb.AppendLine($"/// <summary>Actions of the <c>{Escape(info.ToolName)}</c> tool.</summary>");
        sb.AppendLine($"[JsonConverter(typeof(JsonStringEnumConverter<{enumName}>))]");
        sb.AppendLine($"public enum {enumName}");
        sb.AppendLine("{");

        for (int i = 0; i < info.Methods.Count; i++)
        {
            var method = info.Methods[i];
            if (i > 0)
            {
                sb.AppendLine();
            }

            if (!string.IsNullOrEmpty(method.Summary))
            {
                sb.AppendLine($"    /// <summary>{Escape(method.Summary!)}</summary>");
            }

            sb.Append($"    [JsonStringEnumMemberName(\"{method.ActionName}\")] {method.ActionPascal}");
            sb.AppendLine(i == info.Methods.Count - 1 ? string.Empty : ",");
        }

        sb.AppendLine("}");
        sb.AppendLine();
    }

    private static void EmitTool(
        StringBuilder sb,
        DiscoveredService service,
        List<ExposedParameter> exposed,
        string enumName,
        string className)
    {
        var info = service.Info;
        string methodName = info.CategoryPascal;

        sb.AppendLine("/// <summary>");
        sb.AppendLine($"/// {Escape(info.Title ?? info.ToolName)} for an open Word session.");
        sb.AppendLine("/// </summary>");
        sb.AppendLine("[McpServerToolType]");
        sb.AppendLine($"public static class {className}");
        sb.AppendLine("{");

        sb.AppendLine("    /// <summary>");
        sb.AppendLine($"    /// {Escape(info.Title ?? info.ToolName)}.");
        sb.AppendLine("    /// </summary>");
        sb.AppendLine("    /// <param name=\"action\">The action to perform.</param>");
        sb.AppendLine("    /// <param name=\"session_id\">Session identifier from file(open) or file(create).</param>");

        foreach (var parameter in exposed)
        {
            string doc = Escape(parameter.DescriptionWithRequired ?? parameter.Name);
            sb.AppendLine($"    /// <param name=\"{parameter.WireName}\">{doc}</param>");
        }

        sb.AppendLine("    /// <returns>A JSON payload describing the result.</returns>");

        sb.Append($"    [McpServerTool(Name = \"{info.ToolName}\"");
        if (!string.IsNullOrEmpty(info.Title))
        {
            sb.Append($", Title = \"{Escape(info.Title!)}\"");
        }

        if (info.Destructive == true)
        {
            sb.Append(", Destructive = true");
        }

        sb.AppendLine(")]");

        if (!string.IsNullOrEmpty(info.Description))
        {
            sb.AppendLine($"    [Description({Literal(info.Description!)})]");
        }

        sb.AppendLine($"    public static string {methodName}(");
        sb.AppendLine($"        {enumName} action,");
        sb.Append("        string session_id");

        foreach (var parameter in exposed)
        {
            sb.AppendLine(",");
            EmitParameterDeclaration(sb, parameter);
        }

        sb.AppendLine(")");
        sb.AppendLine($"        => WordToolsBase.Execute(\"{info.ToolName}\", action.ToString(), () =>");
        sb.AppendLine("        {");
        sb.AppendLine("            var batch = WordToolsBase.Batch(session_id);");
        sb.AppendLine();
        sb.AppendLine("            return action switch");
        sb.AppendLine("            {");

        foreach (var method in info.Methods)
        {
            EmitCase(sb, service, method, enumName, exposed);
        }

        sb.AppendLine("                _ => throw new ArgumentException($\"Unknown action '{action}'.\", nameof(action))");
        sb.AppendLine("            };");
        sb.AppendLine("        });");
        sb.AppendLine("}");
    }

    private static void EmitParameterDeclaration(StringBuilder sb, ExposedParameter parameter)
    {
        if (parameter.CanStayNonNullable)
        {
            sb.Append($"        [DefaultValue({parameter.DefaultValue})] "
                + $"{parameter.TypeName} {parameter.WireName} = {parameter.DefaultValue}");
            return;
        }

        string type = parameter.TypeName.EndsWith("?", StringComparison.Ordinal)
            ? parameter.TypeName
            : parameter.TypeName + "?";

        sb.Append($"        [DefaultValue(null)] {type} {parameter.WireName} = null");
    }

    private static void EmitCase(
        StringBuilder sb,
        DiscoveredService service,
        MethodInfo method,
        string enumName,
        List<ExposedParameter> exposed)
    {
        var arguments = new List<string> { "batch" };

        foreach (var parameter in method.Parameters)
        {
            var declared = exposed.Find(p => p.Name == parameter.Name)!;
            arguments.Add(BuildArgument(parameter, declared));
        }

        string call = $"WordServices.{service.ServiceProperty}.{method.MethodName}({string.Join(", ", arguments)})";
        string line = $"                {enumName}.{method.ActionPascal} => {call},";

        if (line.Length <= 118)
        {
            sb.AppendLine(line);
            return;
        }

        sb.AppendLine($"                {enumName}.{method.ActionPascal} => WordServices.{service.ServiceProperty}.{method.MethodName}(");
        for (int i = 0; i < arguments.Count; i++)
        {
            sb.AppendLine($"                    {arguments[i]}{(i == arguments.Count - 1 ? "),": ",")}");
        }
    }

    /// <summary>
    /// Builds the argument expression for one parameter, bridging the gap between the optional
    /// tool parameter and what the command method actually accepts.
    /// </summary>
    private static string BuildArgument(ParameterInfo parameter, ExposedParameter declared)
    {
        string name = declared.WireName;

        // The declared parameter kept its non-nullable type, so it always carries a usable value.
        if (declared.CanStayNonNullable)
        {
            return name;
        }

        // The command accepts null itself, so nothing has to be substituted.
        if (parameter.TypeName.EndsWith("?", StringComparison.Ordinal))
        {
            return name;
        }

        // The command declares a default, so an omitted value falls back to it.
        if (parameter.HasDefault && parameter.DefaultValue != null)
        {
            return $"{name} ?? {parameter.DefaultValue}";
        }

        // The command needs a value and has no fallback, so the omission has to become an error.
        return parameter.TypeName switch
        {
            "string" => $"WordToolsBase.RequireText({name}, nameof({name}))",
            "int" or "double" or "bool" or "long" or "float" or "decimal"
                => $"WordToolsBase.Require({name}, nameof({name}))",
            _ => $"WordToolsBase.RequireValue({name}, nameof({name}))"
        };
    }

    private static string Escape(string text)
        => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    private static string Literal(string text)
        => SyntaxFactory.Literal(text).ToFullString();

    private sealed class DiscoveredService
    {
        public DiscoveredService(ServiceInfo info, string serviceProperty)
        {
            Info = info;
            ServiceProperty = serviceProperty;
        }

        public ServiceInfo Info { get; }

        /// <summary>Name of the WordServices property that returns the command interface.</summary>
        public string ServiceProperty { get; }
    }
}

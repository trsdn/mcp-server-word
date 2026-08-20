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
    private const string ToolAssembly = "WordMcp.McpServer";
    private const string ServiceAssembly = "WordMcp.Service";
    private const string DispatchNamespace = "WordMcp.Service.Generated";

    /// <inheritdoc />
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var services = context.CompilationProvider.Select(static (compilation, _) =>
            (Assembly: compilation.AssemblyName ?? string.Empty, Services: Discover(compilation)));

        context.RegisterSourceOutput(services, static (productionContext, discovered) =>
        {
            // The same description feeds both halves of the boundary: the tool that packs the
            // arguments and the dispatcher that unpacks them. Emitting them from one place is what
            // keeps them from drifting apart.
            switch (discovered.Assembly)
            {
                case ToolAssembly:
                    foreach (var service in discovered.Services)
                    {
                        productionContext.AddSource(
                            $"WordTool.{service.Info.CategoryPascal}.g.cs",
                            SourceText.From(Emit(service.Info), Encoding.UTF8));
                    }

                    break;

                case ServiceAssembly:
                    foreach (var service in discovered.Services)
                    {
                        productionContext.AddSource(
                            $"WordDispatch.{service.Info.CategoryPascal}.g.cs",
                            SourceText.From(EmitDispatch(service), Encoding.UTF8));
                    }

                    productionContext.AddSource(
                        "GeneratedToolDispatch.g.cs",
                        SourceText.From(EmitAggregator(discovered.Services), Encoding.UTF8));
                    break;

                default:
                    break;
            }
        });
    }

    /// <summary>
    /// Finds every command interface of the referenced Core assembly and pairs it with the
    /// WordServices property that exposes it.
    /// </summary>
    private static ImmutableArray<DiscoveredService> Discover(Compilation compilation)
    {
        var documentation = DocumentationSource.Load(compilation, "WordMcp.");
        var builder = ImmutableArray.CreateBuilder<DiscoveredService>();
        var candidates = new List<INamedTypeSymbol>();
        var implementations = new List<INamedTypeSymbol>();

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
                if (type.TypeKind == TypeKind.Interface)
                {
                    candidates.Add(type);
                }
                else if (type is { TypeKind: TypeKind.Class, IsAbstract: false, DeclaredAccessibility: Accessibility.Public })
                {
                    implementations.Add(type);
                }
            }
        }

        foreach (var type in candidates)
        {
            var info = ServiceInfoExtractor.Extract(type, documentation);
            if (info is null)
            {
                continue;
            }

            builder.Add(new DiscoveredService(info, FindImplementation(implementations, type)));
        }

        return builder.ToImmutable();
    }

    /// <summary>
    /// Finds the class implementing a command interface, so the dispatch half can construct it
    /// without a hand-maintained registry that would have to be edited for every new category.
    /// </summary>
    private static string? FindImplementation(List<INamedTypeSymbol> classes, INamedTypeSymbol interfaceSymbol)
    {
        foreach (var candidate in classes)
        {
            foreach (var implemented in candidate.AllInterfaces)
            {
                if (SymbolEqualityComparer.Default.Equals(implemented, interfaceSymbol))
                {
                    return candidate.ToDisplayString();
                }
            }
        }

        return null;
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

    private static string Emit(ServiceInfo info)
    {
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
        EmitTool(sb, info, exposed, enumName, className);

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
        ServiceInfo info,
        List<ExposedParameter> exposed,
        string enumName,
        string className)
    {
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
        sb.AppendLine($"        => WordToolsBase.Execute(\"{info.ToolName}\", action.ToString(), () => ServiceBridge.Invoke(");
        sb.AppendLine("            action switch");
        sb.AppendLine("            {");

        foreach (var method in info.Methods)
        {
            sb.AppendLine($"                {enumName}.{method.ActionPascal} => \"{info.ToolName}.{method.ActionName}\",");
        }

        sb.AppendLine("                _ => throw new ArgumentException($\"Unknown action '{action}'.\", nameof(action))");
        sb.AppendLine("            },");
        sb.AppendLine("            session_id,");

        // The arguments travel as a plain object whose property names are the wire names, so the
        // dispatcher on the other side reads them back under the names the caller typed.
        sb.AppendLine("            new");
        sb.AppendLine("            {");
        for (int i = 0; i < exposed.Count; i++)
        {
            sb.AppendLine($"                {exposed[i].WireName}{(i == exposed.Count - 1 ? string.Empty : ",")}");
        }

        sb.AppendLine("            }));");
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

    /// <summary>
    /// Emits the dispatch half: the class that unpacks one category's arguments and calls the
    /// command implementation. This is the code the tool half talks to across the boundary.
    /// </summary>
    private static string EmitDispatch(DiscoveredService service)
    {
        var info = service.Info;
        var exposed = ServiceInfoExtractor.GetExposedParameters(info);

        var sb = new StringBuilder();
        sb.AppendLine("// <auto-generated/>");
        sb.AppendLine("#nullable enable");
        sb.AppendLine();
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Text.Json;");
        sb.AppendLine("using WordMcp.ComInterop.Session;");
        sb.AppendLine();
        sb.AppendLine($"namespace {DispatchNamespace};");
        sb.AppendLine();
        sb.AppendLine($"/// <summary>Runs the actions of the <c>{Escape(info.ToolName)}</c> tool.</summary>");
        sb.AppendLine($"internal static class {info.CategoryPascal}Dispatch");
        sb.AppendLine("{");

        if (service.Implementation is null)
        {
            // No implementation in reach: emit a dispatcher that declines everything rather than
            // failing the build, so a partially wired category cannot break the whole service.
            sb.AppendLine("    /// <summary>Always declines: no implementation was found.</summary>");
            sb.AppendLine("    internal static bool TryInvoke(string action, Func<IWordBatch> batch, JsonElement args, out object? result)");
            sb.AppendLine("    {");
            sb.AppendLine("        result = null;");
            sb.AppendLine("        return false;");
            sb.AppendLine("    }");
            sb.AppendLine("}");
            return sb.ToString();
        }

        sb.AppendLine($"    private static readonly {info.InterfaceNamespace}.{info.InterfaceName} Commands");
        sb.AppendLine($"        = new {service.Implementation}();");
        sb.AppendLine();
        sb.AppendLine("    /// <summary>Runs one action against an open document.</summary>");
        sb.AppendLine("    /// <param name=\"action\">The action name without the category prefix.</param>");
        sb.AppendLine("    /// <param name=\"batch\">Resolves the session, but only once the action is known.</param>");
        sb.AppendLine("    /// <param name=\"args\">The arguments as the tool packed them.</param>");
        sb.AppendLine("    /// <param name=\"result\">The action's result when it ran.</param>");
        sb.AppendLine("    /// <returns><c>true</c> when the action is known to this category.</returns>");
        sb.AppendLine("    internal static bool TryInvoke(string action, Func<IWordBatch> batch, JsonElement args, out object? result)");
        sb.AppendLine("    {");
        sb.AppendLine("        switch (action)");
        sb.AppendLine("        {");

        foreach (var method in info.Methods)
        {
            var arguments = new List<string> { "batch()" };
            foreach (var parameter in method.Parameters)
            {
                arguments.Add(BuildDispatchArgument(parameter, exposed.Find(p => p.Name == parameter.Name)!));
            }

            sb.AppendLine($"            case \"{method.ActionName}\":");
            sb.AppendLine($"                result = Commands.{method.MethodName}(");
            for (int i = 0; i < arguments.Count; i++)
            {
                sb.AppendLine($"                    {arguments[i]}{(i == arguments.Count - 1 ? ");" : ",")}");
            }

            sb.AppendLine("                return true;");
            sb.AppendLine();
        }

        sb.AppendLine("            default:");
        sb.AppendLine("                result = null;");
        sb.AppendLine("                return false;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine("}");

        return sb.ToString();
    }

    /// <summary>
    /// Emits the one entry point the service calls: it splits <c>category.action</c> and hands the
    /// work to the matching category.
    /// </summary>
    private static string EmitAggregator(ImmutableArray<DiscoveredService> services)
    {
        var sb = new StringBuilder();
        sb.AppendLine("// <auto-generated/>");
        sb.AppendLine("#nullable enable");
        sb.AppendLine();
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Text.Json;");
        sb.AppendLine("using WordMcp.ComInterop.Session;");
        sb.AppendLine();
        sb.AppendLine($"namespace {DispatchNamespace};");
        sb.AppendLine();
        sb.AppendLine("/// <summary>Routes a tool command to the category that implements it.</summary>");
        sb.AppendLine("internal static class GeneratedToolDispatch");
        sb.AppendLine("{");
        sb.AppendLine("    /// <summary>Runs a <c>category.action</c> command against an open document.</summary>");
        sb.AppendLine("    /// <param name=\"command\">The command name, for example <c>text.insert</c>.</param>");
        sb.AppendLine("    /// <param name=\"batch\">Resolves the session, but only once the command is known.</param>");
        sb.AppendLine("    /// <param name=\"args\">The arguments as the tool packed them.</param>");
        sb.AppendLine("    /// <param name=\"result\">The command's result when it ran.</param>");
        sb.AppendLine("    /// <returns><c>true</c> when the command is known.</returns>");
        sb.AppendLine("    internal static bool TryInvoke(string command, Func<IWordBatch> batch, JsonElement args, out object? result)");
        sb.AppendLine("    {");
        sb.AppendLine("        int separator = command.IndexOf('.');");
        sb.AppendLine("        if (separator <= 0)");
        sb.AppendLine("        {");
        sb.AppendLine("            result = null;");
        sb.AppendLine("            return false;");
        sb.AppendLine("        }");
        sb.AppendLine();
        sb.AppendLine("        string action = command.Substring(separator + 1);");
        sb.AppendLine();
        sb.AppendLine("        switch (command.Substring(0, separator))");
        sb.AppendLine("        {");

        foreach (var service in services)
        {
            sb.AppendLine($"            case \"{service.Info.ToolName}\":");
            sb.AppendLine($"                return {service.Info.CategoryPascal}Dispatch.TryInvoke(action, batch, args, out result);");
        }

        sb.AppendLine("            default:");
        sb.AppendLine("                result = null;");
        sb.AppendLine("                return false;");
        sb.AppendLine("        }");
        sb.AppendLine("    }");
        sb.AppendLine("}");

        return sb.ToString();
    }

    /// <summary>
    /// Builds the argument expression on the dispatch side: reads the value back out of the JSON
    /// the tool sent, and reproduces the same fallbacks and the same errors the tool used to raise.
    /// </summary>
    private static string BuildDispatchArgument(ParameterInfo parameter, ExposedParameter declared)
    {
        string name = declared.WireName;
        string bare = parameter.TypeName.TrimEnd('?');
        bool isValueType = bare is "int" or "long" or "double" or "float" or "decimal" or "bool";
        string read = isValueType ? $"ToolArgs.Val<{bare}>(args, \"{name}\")" : $"ToolArgs.Ref<{bare}>(args, \"{name}\")";

        // The tool declared the parameter non-nullable, so every action agrees on a fallback.
        if (declared.CanStayNonNullable)
        {
            return $"{read} ?? {declared.DefaultValue}";
        }

        // The command accepts null itself, so nothing has to be substituted.
        if (parameter.TypeName.EndsWith("?", StringComparison.Ordinal))
        {
            return read;
        }

        // The command declares a default, so an omitted value falls back to it.
        if (parameter.HasDefault && parameter.DefaultValue != null)
        {
            return $"{read} ?? {parameter.DefaultValue}";
        }

        // The command needs a value and has no fallback, so the omission has to become an error.
        return isValueType
            ? $"ToolArgs.RequireVal<{bare}>(args, \"{name}\")"
            : $"ToolArgs.RequireRef<{bare}>(args, \"{name}\")";
    }

    private static string Escape(string text)
        => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    private static string Literal(string text)
        => SyntaxFactory.Literal(text).ToFullString();

    private sealed class DiscoveredService
    {
        public DiscoveredService(ServiceInfo info, string? implementation)
        {
            Info = info;
            Implementation = implementation;
        }

        public ServiceInfo Info { get; }

        /// <summary>Fully qualified name of the class implementing the interface, if one was found.</summary>
        public string? Implementation { get; }
    }
}

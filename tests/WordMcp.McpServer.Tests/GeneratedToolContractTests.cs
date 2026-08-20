using System.ComponentModel;
using System.Reflection;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;
using WordMcp.Core.Attributes;
using Xunit;

namespace WordMcp.McpServer.Tests;

/// <summary>
/// Verifies that the source generator turns every command interface into a matching MCP tool.
/// These tests are the safety net that lets the tool layer be generated instead of written by
/// hand: they compare the emitted surface against the interfaces it is derived from.
/// </summary>
public class GeneratedToolContractTests
{
    /// <summary>Every command interface that is supposed to become a tool.</summary>
    public static TheoryData<Type> CommandInterfaces
    {
        get
        {
            var data = new TheoryData<Type>();
            foreach (var type in FindCommandInterfaces())
            {
                data.Add(type);
            }

            return data;
        }
    }

    [Fact]
    public void Generator_ProducesAToolForEveryCommandInterface()
    {
        var interfaces = FindCommandInterfaces();

        Assert.NotEmpty(interfaces);

        foreach (var type in interfaces)
        {
            Assert.NotNull(FindTool(type));
        }
    }

    [Theory]
    [MemberData(nameof(CommandInterfaces))]
    public void Tool_UsesTheToolNameAndDescriptionFromTheInterface(Type commandInterface)
    {
        var attribute = commandInterface.GetCustomAttribute<McpToolAttribute>()!;
        var method = FindTool(commandInterface)!;

        var toolAttribute = method.GetCustomAttribute<McpServerToolAttribute>();
        Assert.NotNull(toolAttribute);
        Assert.Equal(attribute.ToolName, toolAttribute!.Name);
        Assert.Equal(attribute.Title, toolAttribute.Title);

        // The description is the prompt the model reads, so losing it would silently degrade the
        // server without breaking anything else.
        var description = method.GetCustomAttribute<DescriptionAttribute>();
        Assert.NotNull(description);
        Assert.Equal(attribute.Description, description!.Description);
    }

    [Theory]
    [MemberData(nameof(CommandInterfaces))]
    public void ActionEnum_CoversExactlyTheInterfaceActions(Type commandInterface)
    {
        var expected = commandInterface.GetMethods()
            .Select(m => m.GetCustomAttribute<ServiceActionAttribute>()?.Action)
            .Where(a => a is not null)
            .OrderBy(a => a, StringComparer.Ordinal)
            .ToArray();

        var enumType = FindTool(commandInterface)!.GetParameters()[0].ParameterType;
        Assert.True(enumType.IsEnum, "The first tool parameter has to be the action enum.");

        var actual = enumType.GetFields(BindingFlags.Public | BindingFlags.Static)
            .Select(f => f.GetCustomAttribute<JsonStringEnumMemberNameAttribute>()?.Name)
            .Where(n => n is not null)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(expected, actual);
    }

    [Theory]
    [MemberData(nameof(CommandInterfaces))]
    public void Tool_ExposesEverySessionParameterInSnakeCase(Type commandInterface)
    {
        var method = FindTool(commandInterface)!;
        var exposed = method.GetParameters().Select(p => p.Name).ToHashSet(StringComparer.Ordinal);

        Assert.Contains("session_id", exposed);

        foreach (var command in commandInterface.GetMethods())
        {
            foreach (var parameter in command.GetParameters())
            {
                if (parameter.ParameterType.Name == "IWordBatch")
                {
                    continue;
                }

                string wireName = ToSnakeCase(parameter.Name!);
                Assert.True(
                    exposed.Contains(wireName),
                    $"{method.Name} does not expose '{wireName}' for {commandInterface.Name}.{command.Name}.");
            }
        }
    }

    [Theory]
    [MemberData(nameof(CommandInterfaces))]
    public void Tool_MakesEveryOptionalParameterOmittable(Type commandInterface)
    {
        var parameters = FindTool(commandInterface)!.GetParameters();

        // action and session_id identify the call itself; everything after them belongs to a
        // subset of the actions and therefore has to be optional.
        foreach (var parameter in parameters.Skip(2))
        {
            Assert.True(
                parameter.IsOptional,
                $"{parameter.Name} has to be optional so the other actions can omit it.");
            Assert.NotNull(parameter.GetCustomAttribute<DefaultValueAttribute>());
        }
    }

    private static List<Type> FindCommandInterfaces()
        => typeof(McpToolAttribute).Assembly.GetTypes()
            .Where(t => t.IsInterface
                && t.GetCustomAttribute<ServiceCategoryAttribute>() is not null
                && t.GetCustomAttribute<McpToolAttribute>() is not null)
            .OrderBy(t => t.Name, StringComparer.Ordinal)
            .ToList();

    private static MethodInfo? FindTool(Type commandInterface)
    {
        string toolName = commandInterface.GetCustomAttribute<McpToolAttribute>()!.ToolName;

        return typeof(WordMcp.McpServer.Tools.WordFileTool).Assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .FirstOrDefault(m => m.GetCustomAttribute<McpServerToolAttribute>()?.Name == toolName);
    }

    private static string ToSnakeCase(string camelCase)
    {
        var sb = new System.Text.StringBuilder();
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

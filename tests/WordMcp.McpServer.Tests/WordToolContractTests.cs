using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using WordMcp.McpServer.Tools;
using Xunit;

namespace WordMcp.McpServer.Tests;

/// <summary>
/// Contract tests for the tool layer that need no Word installation.
/// </summary>
public class WordToolContractTests
{
    public static TheoryData<Type> ActionEnums =>
    [
        typeof(WordImageAction),
        typeof(WordFieldAction)
    ];

    [Theory]
    [MemberData(nameof(ActionEnums))]
    public void ActionEnums_ExposeKebabCaseNamesToClients(Type enumType)
    {
        foreach (var member in enumType.GetFields(BindingFlags.Public | BindingFlags.Static))
        {
            var attribute = member.GetCustomAttribute<JsonStringEnumMemberNameAttribute>();

            Assert.True(attribute is not null, $"{enumType.Name}.{member.Name} has no wire name.");
            Assert.Equal(attribute!.Name, attribute.Name.ToLowerInvariant());
            Assert.DoesNotContain(" ", attribute.Name, StringComparison.Ordinal);
            Assert.DoesNotContain("_", attribute.Name, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Image_ReportsUnknownSessionAsStructuredError()
    {
        var json = WordImageTool.Image(WordImageAction.List, "word-does-not-exist");

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("image", root.GetProperty("tool").GetString());
        Assert.Equal("List", root.GetProperty("action").GetString());
    }

    [Fact]
    public void Field_ReportsUnknownSessionAsStructuredError()
    {
        var json = WordFieldTool.Field(WordFieldAction.List, "word-does-not-exist");

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("field", root.GetProperty("tool").GetString());
    }

    [Fact]
    public void Field_ValidatesArgumentsBeforeTouchingTheSession()
    {
        // A bad position must be reported as such, not swallowed by the missing session.
        var json = WordFieldTool.Field(
            WordFieldAction.InsertToc, "word-does-not-exist", upper_heading_level: 5, lower_heading_level: 1);

        using var doc = JsonDocument.Parse(json);
        Assert.False(doc.RootElement.GetProperty("success").GetBoolean());
    }
}

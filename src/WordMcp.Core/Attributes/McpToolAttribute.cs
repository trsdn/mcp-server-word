namespace WordMcp.Core.Attributes;

/// <summary>
/// Describes how a command interface is surfaced as an MCP tool.
/// </summary>
[AttributeUsage(AttributeTargets.Interface | AttributeTargets.Method)]
public sealed class McpToolAttribute : Attribute
{
    /// <summary>
    /// Creates the attribute.
    /// </summary>
    /// <param name="toolName">MCP tool name, for example <c>document</c>.</param>
    public McpToolAttribute(string toolName)
        => ToolName = toolName ?? throw new ArgumentNullException(nameof(toolName));

    /// <summary>Gets the MCP tool name.</summary>
    public string ToolName { get; }

    /// <summary>Gets or sets the human-readable tool title.</summary>
    public string? Title { get; set; }

    /// <summary>Gets or sets the tool description shown to the model.</summary>
    public string? Description { get; set; }

    /// <summary>Gets or sets a value indicating whether the tool can modify the document.</summary>
    public bool Destructive { get; set; } = true;
}

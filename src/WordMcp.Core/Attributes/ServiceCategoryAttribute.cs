namespace WordMcp.Core.Attributes;

/// <summary>
/// Marks a command interface as a service category. The category name is the stable identifier
/// used by MCP tool names, CLI commands and generated skill documentation.
/// </summary>
/// <remarks>
/// Applied to interfaces such as <c>IDocumentCommands</c>. A future source generator reads this
/// attribute to emit MCP tool and CLI command wrappers, mirroring the PptMcp code generation model.
/// </remarks>
[AttributeUsage(AttributeTargets.Interface)]
public sealed class ServiceCategoryAttribute : Attribute
{
    /// <summary>
    /// Creates the attribute.
    /// </summary>
    /// <param name="category">Lower-case category name, for example <c>document</c>.</param>
    /// <param name="pascalName">Optional PascalCase form used for generated type names.</param>
    public ServiceCategoryAttribute(string category, string? pascalName = null)
    {
        Category = category ?? throw new ArgumentNullException(nameof(category));
        PascalName = pascalName;
    }

    /// <summary>Gets the lower-case category name.</summary>
    public string Category { get; }

    /// <summary>Gets the PascalCase form used for generated type names, if provided.</summary>
    public string? PascalName { get; }
}

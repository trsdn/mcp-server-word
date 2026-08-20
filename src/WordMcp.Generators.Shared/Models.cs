namespace WordMcp.Generators.Common;

/// <summary>
/// A command interface marked with [ServiceCategory], flattened into everything the generators need.
/// </summary>
public sealed class ServiceInfo
{
    public ServiceInfo(
        string category,
        string categoryPascal,
        string toolName,
        string? title,
        string? description,
        bool? destructive,
        List<MethodInfo> methods,
        string interfaceName,
        string interfaceNamespace)
    {
        Category = category;
        CategoryPascal = categoryPascal;
        ToolName = toolName;
        Title = title;
        Description = description;
        Destructive = destructive;
        Methods = methods;
        InterfaceName = interfaceName;
        InterfaceNamespace = interfaceNamespace;
    }

    /// <summary>Kebab-case category name, for example <c>header-footer</c>.</summary>
    public string Category { get; }

    /// <summary>PascalCase form used for generated type names, for example <c>HeaderFooter</c>.</summary>
    public string CategoryPascal { get; }

    /// <summary>The MCP tool name clients call.</summary>
    public string ToolName { get; }

    public string? Title { get; }

    /// <summary>The text the language model sees. Comes from [McpTool(Description = ...)].</summary>
    public string? Description { get; }

    public bool? Destructive { get; }

    public List<MethodInfo> Methods { get; }

    public string InterfaceName { get; }

    public string InterfaceNamespace { get; }
}

/// <summary>
/// One action of a service category.
/// </summary>
public sealed class MethodInfo
{
    public MethodInfo(string methodName, string actionName, string actionPascal,
        List<ParameterInfo> parameters, string? summary)
    {
        MethodName = methodName;
        ActionName = actionName;
        ActionPascal = actionPascal;
        Parameters = parameters;
        Summary = summary;
    }

    /// <summary>The C# method name on the interface.</summary>
    public string MethodName { get; }

    /// <summary>Kebab-case wire name clients pass as the action.</summary>
    public string ActionName { get; }

    /// <summary>PascalCase name of the generated enum member.</summary>
    public string ActionPascal { get; }

    /// <summary>Parameters without the leading IWordBatch.</summary>
    public List<ParameterInfo> Parameters { get; }

    public string? Summary { get; }
}

/// <summary>
/// A single parameter of an action.
/// </summary>
public sealed class ParameterInfo
{
    public ParameterInfo(string name, string typeName, bool hasDefault, string? defaultValue,
        string? description)
    {
        Name = name;
        TypeName = typeName;
        HasDefault = hasDefault;
        DefaultValue = defaultValue;
        Description = description;
    }

    /// <summary>camelCase name as declared on the interface.</summary>
    public string Name { get; }

    /// <summary>Fully qualified type name, with nullable annotation preserved.</summary>
    public string TypeName { get; }

    public bool HasDefault { get; }

    /// <summary>The default value as a C# expression, or <c>null</c> when there is none.</summary>
    public string? DefaultValue { get; }

    public string? Description { get; }

    /// <summary>
    /// A parameter the caller must supply for this action: no default and not nullable.
    /// </summary>
    public bool IsRequired => !HasDefault && !TypeName.EndsWith("?", StringComparison.Ordinal);
}

/// <summary>
/// A parameter as it appears on the generated tool method, merged across all actions of the
/// category. The tool has one method for every action, so a parameter used by only some actions
/// still has to be declared — and therefore has to be optional.
/// </summary>
public sealed class ExposedParameter
{
    public ExposedParameter(string name, string typeName, string? description)
    {
        Name = name;
        TypeName = typeName;
        Description = description;
    }

    /// <summary>camelCase name, matching the interface parameter it came from.</summary>
    public string Name { get; }

    /// <summary>snake_case name as it appears on the wire.</summary>
    public string WireName => NameHelper.ToSnakeCase(Name);

    /// <summary>The merged type. Nullable as soon as any action declares it nullable.</summary>
    public string TypeName { get; set; }

    public string? Description { get; set; }

    /// <summary>
    /// The shared default of every action that declares this parameter, or <c>null</c> when the
    /// actions disagree or at least one action leaves it without a default.
    /// </summary>
    public string? DefaultValue { get; set; }

    /// <summary>Set once the actions disagree about the default, which forces a nullable parameter.</summary>
    public bool DefaultsConflict { get; set; }

    /// <summary>
    /// Whether the generated parameter can keep its non-nullable type. That needs a default every
    /// action agrees on, and no action that treats the parameter as mandatory.
    /// </summary>
    public bool CanStayNonNullable
        => !TypeName.EndsWith("?", StringComparison.Ordinal)
            && !DefaultsConflict
            && DefaultValue != null
            && RequiredByActions.Count == 0;

    /// <summary>Actions that cannot run without this parameter.</summary>
    public List<string> RequiredByActions { get; } = new();

    /// <summary>Total number of actions, used to phrase the "required" hint.</summary>
    public int TotalActionCount { get; set; }

    /// <summary>
    /// The description with a hint about which actions need the parameter, so the model can tell
    /// a globally required parameter from one that only matters for some actions.
    /// </summary>
    public string? DescriptionWithRequired
    {
        get
        {
            if (RequiredByActions.Count == 0)
            {
                return Description;
            }

            var suffix = RequiredByActions.Count == TotalActionCount
                ? "(required)"
                : $"(required for: {string.Join(", ", RequiredByActions)})";

            return string.IsNullOrEmpty(Description) ? suffix : $"{Description} {suffix}";
        }
    }
}

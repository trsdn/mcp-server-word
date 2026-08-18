namespace WordMcp.Core.Attributes;

/// <summary>
/// Marks a command method as an action of its service category. The action name is the value
/// callers pass as the <c>action</c> parameter of the corresponding MCP tool.
/// </summary>
[AttributeUsage(AttributeTargets.Method)]
public sealed class ServiceActionAttribute : Attribute
{
    /// <summary>
    /// Creates the attribute.
    /// </summary>
    /// <param name="action">Lower-case action name, for example <c>get-info</c>.</param>
    public ServiceActionAttribute(string action)
        => Action = action ?? throw new ArgumentNullException(nameof(action));

    /// <summary>Gets the action name.</summary>
    public string Action { get; }
}

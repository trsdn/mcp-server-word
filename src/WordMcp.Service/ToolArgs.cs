using System.Text.Json;

namespace WordMcp.Service;

/// <summary>
/// Reads tool arguments back out of the JSON object the MCP server sent.
/// </summary>
/// <remarks>
/// The tool layer packs its parameters under their wire names and drops the ones the caller
/// omitted, so every read here has to cope with an absent property. The error messages repeat
/// what the tool layer used to say word for word: the caller names a parameter, not a transport,
/// and should not be able to tell which side of the process boundary noticed the omission.
/// </remarks>
public static class ToolArgs
{
    /// <summary>
    /// Reads an optional value-typed argument.
    /// </summary>
    /// <typeparam name="T">The value type to read.</typeparam>
    /// <param name="args">The argument object.</param>
    /// <param name="name">The wire name of the argument.</param>
    /// <returns>The value, or <c>null</c> when the caller omitted it.</returns>
    public static T? Val<T>(JsonElement args, string name)
        where T : struct
        => TryGet(args, name, out var property) ? property.Deserialize<T>(ServiceProtocol.JsonOptions) : null;

    /// <summary>
    /// Reads an optional reference-typed argument.
    /// </summary>
    /// <typeparam name="T">The reference type to read.</typeparam>
    /// <param name="args">The argument object.</param>
    /// <param name="name">The wire name of the argument.</param>
    /// <returns>The value, or <c>null</c> when the caller omitted it.</returns>
    public static T? Ref<T>(JsonElement args, string name)
        where T : class
        => TryGet(args, name, out var property) ? property.Deserialize<T>(ServiceProtocol.JsonOptions) : null;

    /// <summary>
    /// Reads a value-typed argument the action cannot run without.
    /// </summary>
    /// <typeparam name="T">The value type to read.</typeparam>
    /// <param name="args">The argument object.</param>
    /// <param name="name">The wire name of the argument.</param>
    /// <returns>The value.</returns>
    /// <exception cref="ArgumentException">The caller omitted the argument.</exception>
    public static T RequireVal<T>(JsonElement args, string name)
        where T : struct
        => Val<T>(args, name) ?? throw Missing(name);

    /// <summary>
    /// Reads a reference-typed argument the action cannot run without. A blank string counts as
    /// missing, because a whitespace-only style or bookmark name is never what the caller meant.
    /// </summary>
    /// <typeparam name="T">The reference type to read.</typeparam>
    /// <param name="args">The argument object.</param>
    /// <param name="name">The wire name of the argument.</param>
    /// <returns>The value.</returns>
    /// <exception cref="ArgumentException">The caller omitted the argument.</exception>
    public static T RequireRef<T>(JsonElement args, string name)
        where T : class
    {
        var value = Ref<T>(args, name);
        if (value is null || (value is string text && string.IsNullOrWhiteSpace(text)))
        {
            throw Missing(name);
        }

        return value;
    }

    private static bool TryGet(JsonElement args, string name, out JsonElement property)
    {
        if (args.ValueKind != JsonValueKind.Object
            || !args.TryGetProperty(name, out property)
            || property.ValueKind == JsonValueKind.Null)
        {
            property = default;
            return false;
        }

        return true;
    }

    private static ArgumentException Missing(string name)
        => new($"{name} is required for this action.", name);
}

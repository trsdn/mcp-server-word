using System.Reflection;
using System.Runtime.Loader;

namespace WordMcp.ComInterop;

/// <summary>
/// Resolves the Office interop assemblies from the application directory.
/// </summary>
/// <remarks>
/// <para>Word is driven through <c>dynamic</c>, and the runtime binder needs the real interop
/// assemblies to bind a member — including <c>office.dll</c>, which
/// <c>Microsoft.Office.Interop.Word</c> refers to but which no package declares as a runtime
/// dependency. The file is copied next to every executable, but a framework-dependent app only
/// resolves what its deps file lists, so the load fails with a <see cref="FileNotFoundException"/>
/// the moment the first dynamic call is bound.</para>
/// <para>A test host hides this, because it probes its own directory. A plain executable does not,
/// which is why the session service could not open a single document until this existed.</para>
/// </remarks>
internal static class ComAssemblyResolver
{
    private static readonly string[] Prefixes =
    [
        "office",
        "stdole",
        "Microsoft.Office.",
        "Microsoft.Vbe."
    ];

    private static int _installed;

    /// <summary>
    /// Installs the resolver once. Called by the types that talk to Word, so no host has to know
    /// this is necessary — forgetting it is exactly the bug this fixes.
    /// </summary>
    internal static void Install()
    {
        if (Interlocked.Exchange(ref _installed, 1) == 0)
        {
            AssemblyLoadContext.Default.Resolving += Resolve;
        }
    }

    private static Assembly? Resolve(AssemblyLoadContext context, AssemblyName name)
    {
        // Deliberately narrow: this exists for the Office interop assemblies, and silently
        // loading anything else that happens to sit in the directory would hide real problems.
        if (name.Name is null || !Prefixes.Any(p => name.Name.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
        {
            return null;
        }

        var candidate = Path.Combine(AppContext.BaseDirectory, name.Name + ".dll");
        return File.Exists(candidate) ? context.LoadFromAssemblyPath(candidate) : null;
    }
}

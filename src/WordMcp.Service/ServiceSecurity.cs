using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;

namespace WordMcp.Service;

/// <summary>
/// Names and secures the named pipe the session service listens on.
/// </summary>
/// <remarks>
/// <para>Two layers keep one user's documents away from another's. The pipe name embeds the
/// caller's SID, so two users never resolve the same pipe in the first place; and the pipe is
/// created with an ACL that grants full control to that SID alone, so guessing the name is not
/// enough either. Named pipes on <c>.</c> are local IPC, so nothing here is reachable from the
/// network.</para>
/// <para>Clients additionally pass <see cref="PipeOptions.CurrentUserOnly"/>, which makes
/// Windows verify that the server was created by the same user before any bytes are exchanged.
/// That closes the reverse direction: a rogue process cannot impersonate the service.</para>
/// </remarks>
public static class ServiceSecurity
{
    private static readonly Lazy<string> LazyUserSid = new(() =>
    {
        var sid = WindowsIdentity.GetCurrent().User?.Value;
        return string.IsNullOrEmpty(sid)
            ? throw new InvalidOperationException(
                "Cannot determine the current user SID. Named pipe isolation requires one.")
            : sid;
    });

    /// <summary>
    /// Gets the SID of the current user.
    /// </summary>
    public static string UserSid => LazyUserSid.Value;

    /// <summary>
    /// Gets the name of the shared per-user session service pipe.
    /// </summary>
    /// <returns>The pipe name.</returns>
    public static string GetServicePipeName() => $"WordMcp-service-{UserSid}";

    /// <summary>
    /// Builds a pipe name that is unique to one process, for tests and for private instances.
    /// </summary>
    /// <param name="instance">A discriminator that makes the name unique.</param>
    /// <returns>The pipe name.</returns>
    /// <exception cref="ArgumentException"><paramref name="instance"/> is blank.</exception>
    public static string GetPrivatePipeName(string instance)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(instance);
        return $"WordMcp-private-{UserSid}-{instance}";
    }

    /// <summary>
    /// Creates a listening pipe that only the current user may open.
    /// </summary>
    /// <param name="pipeName">Name of the pipe.</param>
    /// <returns>The pipe server stream, ready to await a connection.</returns>
    /// <exception cref="ArgumentException"><paramref name="pipeName"/> is blank.</exception>
    public static NamedPipeServerStream CreateSecureServer(string pipeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pipeName);

        var security = new PipeSecurity();
        security.AddAccessRule(new PipeAccessRule(
            WindowsIdentity.GetCurrent().User!,
            PipeAccessRights.FullControl,
            AccessControlType.Allow));

        return NamedPipeServerStreamAcl.Create(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: NamedPipeServerStream.MaxAllowedServerInstances,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            inBufferSize: 4096,
            outBufferSize: 4096,
            security);
    }

    /// <summary>
    /// Creates a client for a service pipe.
    /// </summary>
    /// <param name="pipeName">Name of the pipe.</param>
    /// <returns>An unconnected client stream.</returns>
    /// <exception cref="ArgumentException"><paramref name="pipeName"/> is blank.</exception>
    public static NamedPipeClientStream CreateClient(string pipeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pipeName);

        return new NamedPipeClientStream(
            ".",
            pipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
    }
}

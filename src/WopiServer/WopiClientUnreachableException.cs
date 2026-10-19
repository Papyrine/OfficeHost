namespace WopiServer;

/// <summary>
/// No reachable WOPI client offers the action: the discovery of some could not be fetched, as when a WOPI client is down.
/// </summary>
public class WopiClientUnreachableException(string message, IReadOnlyList<string> discoveryUrls) :
    InvalidOperationException(message)
{
    /// <summary>
    /// The discovery URLs that could not be fetched.
    /// </summary>
    public IReadOnlyList<string> DiscoveryUrls { get; } = discoveryUrls;
}
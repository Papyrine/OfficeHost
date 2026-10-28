using System;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

/// <summary>
/// Which documents the module copies: a percentage of them, by a hash (FNV-1a) of their URL, the same on every server
/// and in every app domain, so that a document's exchanges are all copied or none (the shadow needs them all).
/// </summary>
static class TapSampler
{
    static Regex url = new(
        "<(?:\\w+:)?Request\\s[^>]*?\\bUrl=\"([^\"]+)\"",
        RegexOptions.CultureInvariant,
        TimeSpan.FromMilliseconds(100));

    /// <summary>
    /// The URL of the document a cell storage request names: the Url of its first Request, in the SOAP envelope at the
    /// start of the body.
    /// </summary>
    public static string? Document(string start)
    {
        try
        {
            var match = url.Match(start);
            if (!match.Success)
            {
                return null;
            }

            return WebUtility.HtmlDecode(match.Groups[1].Value);
        }
        catch (RegexMatchTimeoutException)
        {
            return null;
        }
    }

    /// <summary>
    /// What tells documents apart: the host and the path, unescaped, in lower case, as the shadow's ShadowDocument.Key.
    /// </summary>
    public static string Key(string document)
    {
        if (!Uri.TryCreate(document, UriKind.Absolute, out var uri))
        {
            return document.ToLowerInvariant();
        }

        return $"{uri.Host}/{Uri.UnescapeDataString(uri.AbsolutePath).Trim('/')}".ToLowerInvariant();
    }

    public static bool Sampled(string key, int percent)
    {
        if (percent >= 100)
        {
            return true;
        }

        if (percent <= 0)
        {
            return false;
        }

        return Hash(key) % 100 < percent;
    }

    static uint Hash(string value)
    {
        var hash = 2166136261;
        foreach (var value8 in Encoding.UTF8.GetBytes(value))
        {
            hash ^= value8;
            hash *= 16777619;
        }

        return hash;
    }
}

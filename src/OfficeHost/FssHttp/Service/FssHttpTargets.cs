/// <summary>
/// The file a document URL of a desktop site names, <c>{site}/{name}</c>, as SharePoint finds a document by its URL: the
/// file of the token under its current name, or a file the site made by Save As (<see cref="SaveAsFiles"/>), or, for a
/// Save As, a name in the folder of the file that the site can create or take over. An old name, after a rename, names
/// nothing, as on SharePoint Server SE (docs/fsshttp-notes.md).
/// </summary>
class FssHttpTargets(IOfficeStorage storage, SaveAsFiles saveAs)
{
    /// <summary>
    /// The name a URL gives after the site of the token, or null for a URL that does not have the site in its path (the
    /// file of the token).
    /// </summary>
    public static string? NameOf(FssHttpSite site, string? url)
    {
        if (url == null ||
            !Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return null;
        }

        var path = uri.AbsolutePath;
        var marker = $"/{site.Token}/";
        var index = path.IndexOf(marker, StringComparison.Ordinal);
        if (index < 0)
        {
            return null;
        }

        var name = Uri.UnescapeDataString(path[(index + marker.Length)..]);
        if (name.Length == 0)
        {
            return null;
        }

        return name;
    }

    public async Task<FssHttpTarget> Find(FssHttpSite site, string? name, Cancel cancel = default)
    {
        if (name == null)
        {
            return new(site);
        }

        var file = await storage.GetFile(site.FileId, cancel);
        if (file == null ||
            string.Equals(name, file.Name, StringComparison.OrdinalIgnoreCase))
        {
            return new(site);
        }

        if (!FileNames.IsLegal(name))
        {
            return new(null);
        }

        var sibling = await storage.GetSibling(site.FileId, name, cancel);
        if (sibling != null &&
            (await saveAs.Get(site, cancel)).Contains(sibling.Id))
        {
            return new(site.For(sibling.Id));
        }

        if (!CanSaveAs(site, name))
        {
            return new(null);
        }

        return new(null, name, sibling);
    }

    /// <summary>
    /// The site of a file by its ResourceID, as IsOnlyClient names it: the file of the token or one the site made by Save
    /// As, or null.
    /// </summary>
    public async Task<FssHttpSite?> Find(FssHttpSite site, Guid resourceId, Cancel cancel = default)
    {
        if (resourceId == site.ResourceId)
        {
            return site;
        }

        foreach (var fileId in await saveAs.Get(site, cancel))
        {
            if (FssHttpSite.ResourceIdOf(fileId) == resourceId)
            {
                return site.For(fileId);
            }
        }

        return null;
    }

    /// <summary>
    /// Whether the site can save as a file of the name: with <see cref="OfficeAccess.CanWriteRelative"/>, as for the Save
    /// As of the WOPI clients, and in one of the formats desktop Office edits through OfficeHost.
    /// </summary>
    static bool CanSaveAs(FssHttpSite site, string name) =>
        site.Access is
        {
            CanWrite: true,
            CanWriteRelative: true
        } &&
        OfficeLauncher.DesktopScheme(FileNames.Extension(name).ToLowerInvariant()) != null;
}

/// <summary>
/// What a document URL of a site names: the site of an existing file it reaches, or else, for a Save As, the name of a
/// file to create, and the file of that name that a Save As would take over.
/// </summary>
record FssHttpTarget(FssHttpSite? Site, string? SaveAsName = null, OfficeFile? SaveAsExisting = null);

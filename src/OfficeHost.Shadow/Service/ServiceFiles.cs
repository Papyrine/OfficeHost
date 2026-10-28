/// <summary>
/// The files of the documents in the application's storage (<see cref="IShadowDocuments"/>), each with a token for the
/// user of OfficeHost of a user of SharePoint (<see cref="OfficeHostShadowOptions.Users"/>), issued again before half
/// its lifetime has passed.
/// </summary>
sealed class ServiceFiles(
    IShadowDocuments documents,
    IOfficeTokens tokens,
    IOptions<OfficeHostOptions> host,
    IOptions<OfficeHostShadowOptions> options) :
    IShadowFiles
{
    Dictionary<string, string> ids = [];
    Dictionary<(string FileId, string User), OfficeToken> issued = [];

    public async Task<ShadowFile?> Open(ShadowDocument document, string user, Cancel cancel)
    {
        if (!ids.TryGetValue(document.Key, out var fileId))
        {
            var found = await documents.Find(document, cancel);
            if (found == null)
            {
                return null;
            }

            fileId = found;
            ids[document.Key] = fileId;
        }

        var shadowUser = options.Value.Users(user);
        var now = DateTimeOffset.UtcNow;
        if (!issued.TryGetValue((fileId, shadowUser.Id), out var token) ||
            token.Expires - now < host.Value.TokenLifetime / 2)
        {
            token = await tokens.Issue(
                new(fileId, shadowUser.Id)
                {
                    UserName = shadowUser.Name,
                    CanWrite = true,
                    CanWriteRelative = true,
                    CanRename = true
                },
                cancel);
            issued[(fileId, shadowUser.Id)] = token;
        }

        return new(fileId, $"{host.Value.DesktopPath.TrimEnd('/')}/{token.Value}");
    }

    /// <summary>
    /// Forgets the file of a document and its tokens, as when it is written again.
    /// </summary>
    public void Forget(string key)
    {
        if (!ids.Remove(key, out var fileId))
        {
            return;
        }

        foreach (var entry in issued.Keys.Where(_ => _.FileId == fileId).ToList())
        {
            issued.Remove(entry);
        }
    }
}

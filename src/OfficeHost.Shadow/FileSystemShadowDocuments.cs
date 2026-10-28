namespace OfficeHost;

/// <summary>
/// The documents of SharePoint in a <see cref="FileSystemStorage"/>: each at its host and path under the root of the
/// storage (<c>{root}/sharepoint/sites/team/Shared Documents/Budget.xlsx</c>).
/// </summary>
public class FileSystemShadowDocuments(FileSystemStorage storage) :
    IShadowDocuments
{
    public Task<string?> Find(ShadowDocument document, Cancel cancel = default)
    {
        var path = PathOf(document);
        if (!File.Exists(Path.Combine(storage.Root, path)))
        {
            return Task.FromResult<string?>(null);
        }

        return Task.FromResult<string?>(storage.GetFileId(path));
    }

    public async Task<string> Add(ShadowDocument document, Stream content, Cancel cancel = default)
    {
        var path = PathOf(document);
        var full = Path.Combine(storage.Root, path);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        await using (var file = File.Create(full))
        {
            await content.CopyToAsync(file, cancel);
        }

        return storage.GetFileId(path);
    }

    static string PathOf(ShadowDocument document)
    {
        string[] segments = [document.Url.Host, .. document.Path.Split('/')];
        // a path of the traffic, kept under the root
        if (segments.Any(_ => _.Length == 0 || _ is "." or ".." || _.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0))
        {
            throw new InvalidDataException($"Not a path of a document: {document.Url}");
        }

        return Path.Combine(segments);
    }
}

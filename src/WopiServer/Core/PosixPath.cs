/// <summary>
/// Python <c>os.path</c> semantics on POSIX for storage paths. Storage paths always use `/`,
/// whatever the host OS, so <see cref="Path"/> (which uses `\` on Windows) must not be used on them.
/// </summary>
static class PosixPath
{
    public const string Sep = "/";

    public static string DirName(string path)
    {
        var index = path.LastIndexOf('/') + 1;
        var head = path[..index];
        if (head.Length > 0 &&
            head.Any(_ => _ != '/'))
        {
            head = head.TrimEnd('/');
        }

        return head;
    }

    public static string BaseName(string path) =>
        path[(path.LastIndexOf('/') + 1)..];

    /// <summary>
    /// Splits the extension, including its leading dot, off the path. As in Python, leading dots
    /// of the file name do not start an extension: `.bashrc` has none.
    /// </summary>
    public static (string Root, string Extension) SplitExt(string path)
    {
        var sepIndex = path.LastIndexOf('/');
        var dotIndex = path.LastIndexOf('.');
        if (dotIndex > sepIndex)
        {
            for (var index = sepIndex + 1; index < dotIndex; index++)
            {
                if (path[index] != '.')
                {
                    return (path[..dotIndex], path[dotIndex..]);
                }
            }
        }

        return (path, "");
    }

    public static string Extension(string path) =>
        SplitExt(path).Extension;

    /// <summary>
    /// Joins a folder and a name with a single separator, so that files in the root folder do not get a `//` prefix.
    /// </summary>
    public static string Join(string folder, string name)
    {
        if (folder.EndsWith('/'))
        {
            return $"{folder}{name}";
        }

        return $"{folder}/{name}";
    }
}

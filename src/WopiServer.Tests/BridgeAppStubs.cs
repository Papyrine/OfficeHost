using System.IO.Compression;
using Microsoft.AspNetCore.WebUtilities;

/// <summary>
/// The CodiMD endpoints used by the bridge, holding the notes and uploads in memory.
/// </summary>
sealed class CodiMDStub
{
    public Dictionary<string, string> Notes { get; } = [];
    public Dictionary<string, byte[]> Uploads { get; } = [];

    /// <summary>
    /// A docid that CodiMD redirects to another one, as for an aliased note.
    /// </summary>
    public (string From, string To)? Alias { get; set; }

    public bool Available { get; set; } = true;
    public bool DownloadFails { get; set; }

    public async Task<HttpResponseMessage> Handle(HttpRequestMessage request)
    {
        var uri = request.RequestUri!;
        var path = Uri.UnescapeDataString(uri.AbsolutePath);
        if (!Available)
        {
            return new(HttpStatusCode.NotFound);
        }

        if (request.Method == HttpMethod.Head)
        {
            if (path.StartsWith("/uploads/", StringComparison.Ordinal))
            {
                if (!Uploads.TryGetValue(path["/uploads/".Length..], out var upload))
                {
                    return new(HttpStatusCode.NotFound);
                }

                var head = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent([])
                };
                head.Content.Headers.ContentLength = upload.Length;
                return head;
            }

            if (Alias is { } alias &&
                path == $"/{alias.From}")
            {
                var redirect = new HttpResponseMessage(HttpStatusCode.Found);
                redirect.Headers.Location = new($"http://codimd/{alias.To}");
                return redirect;
            }

            return new(HttpStatusCode.OK);
        }

        if (request.Method == HttpMethod.Post &&
            path == "/new")
        {
            var id = $"new{Notes.Count}";
            Notes[id] = await request.Content!.ReadAsStringAsync();
            var created = new HttpResponseMessage(HttpStatusCode.Found);
            created.Headers.Location = new($"http://codimd/{id}");
            return created;
        }

        if (request.Method == HttpMethod.Post &&
            path == "/uploadimage")
        {
            var form = (MultipartFormDataContent) request.Content!;
            var file = form.Single();
            Uploads[file.Headers.ContentDisposition!.FileName!.Trim('"')] = await file.ReadAsByteArrayAsync();
            return new(HttpStatusCode.OK);
        }

        if (request.Method == HttpMethod.Put &&
            path.StartsWith("/api/notes/", StringComparison.Ordinal))
        {
            var json = JsonNode.Parse(await request.Content!.ReadAsStringAsync())!;
            Notes[path["/api/notes/".Length..]] = (string) json["content"]!;
            return new(HttpStatusCode.OK);
        }

        if (request.Method == HttpMethod.Get &&
            path.StartsWith("/uploads/", StringComparison.Ordinal) &&
            Uploads.TryGetValue(path["/uploads/".Length..], out var content))
        {
            return new(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(content)
            };
        }

        if (request.Method == HttpMethod.Get &&
            path.EndsWith("/download", StringComparison.Ordinal))
        {
            if (DownloadFails)
            {
                return new(HttpStatusCode.InternalServerError)
                {
                    Content = new StringContent("down")
                };
            }

            return new(HttpStatusCode.OK)
            {
                Content = new StringContent(Notes[path[1..^"/download".Length]])
            };
        }

        return new(HttpStatusCode.NotFound);
    }
}

/// <summary>
/// The Etherpad API and endpoints used by the bridge, holding the pads in memory.
/// </summary>
sealed class EtherpadStub
{
    public Dictionary<string, string> Pads { get; } = [];

    /// <summary>
    /// The save webhook metadata given to Etherpad, by pad.
    /// </summary>
    public Dictionary<string, (string WopiSrc, string AccessToken)> Metadata { get; } = [];

    public List<string> Calls { get; } = [];

    public async Task<HttpResponseMessage> Handle(HttpRequestMessage request)
    {
        var uri = request.RequestUri!;
        var path = uri.AbsolutePath;
        var query = QueryHelpers.ParseQuery(uri.Query);
        if (request.Method == HttpMethod.Post &&
            path.StartsWith("/api/1/", StringComparison.Ordinal))
        {
            var method = path["/api/1/".Length..];
            Calls.Add(method);
            if (query["apikey"] != "etherpadkey")
            {
                return Json("""{"code": 4, "message": "no or wrong API Key"}""");
            }

            return method switch
            {
                "createGroupIfNotExistsFor" => Json("""{"code": 0, "message": "ok", "data": {"groupID": "g.group"}}"""),
                "getReadOnlyID" => Json(new JsonObject { ["code"] = 0, ["message"] = "ok", ["data"] = new JsonObject { ["readOnlyID"] = $"r.{query["padID"]}" } }.ToJsonString()),
                _ => Json("""{"code": 0, "message": "ok", "data": null}""")
            };
        }

        if (request.Method == HttpMethod.Post &&
            path == "/setEFSSMetadata")
        {
            Metadata[query["padID"].ToString()] = (Uri.UnescapeDataString(query["wopiSrc"].ToString()), query["accessToken"].ToString());
            return Json("""{"code": 0}""");
        }

        if (request.Method == HttpMethod.Post &&
            path.EndsWith("/import", StringComparison.Ordinal))
        {
            var form = (MultipartFormDataContent) request.Content!;
            Pads[path["/p/".Length..^"/import".Length]] = await form.Single().ReadAsStringAsync();
            return new(HttpStatusCode.OK);
        }

        if (request.Method == HttpMethod.Get &&
            path.EndsWith("/export/etherpad", StringComparison.Ordinal))
        {
            return new(HttpStatusCode.OK)
            {
                Content = new StringContent(Pads[path["/p/".Length..^"/export/etherpad".Length]])
            };
        }

        return new(HttpStatusCode.NotFound);
    }

    static HttpResponseMessage Json(string json) =>
        new(HttpStatusCode.OK)
        {
            Content = new StringContent(json)
        };
}

/// <summary>
/// Zip bundles of CodiMD (.zmd): a markdown file plus its pictures.
/// </summary>
static class Bundle
{
    public static byte[] Create(params (string Name, byte[] Content)[] entries)
    {
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, true))
        {
            foreach (var (name, content) in entries)
            {
                using var stream = zip.CreateEntry(name).Open();
                stream.Write(content);
            }
        }

        return buffer.ToArray();
    }

    public static Dictionary<string, string> Read(byte[] bundle)
    {
        using var zip = new ZipArchive(new MemoryStream(bundle), ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(
            _ => _.FullName,
            _ =>
            {
                using var reader = new StreamReader(_.Open());
                return reader.ReadToEnd();
            });
    }
}

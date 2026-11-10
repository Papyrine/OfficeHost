/// <summary>
/// A minimal Word document.
/// </summary>
static class Docx
{
    public static byte[] Create(string text)
    {
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, true))
        {
            Add(
                zip,
                "[Content_Types].xml",
                """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""");
            Add(
                zip,
                "_rels/.rels",
                """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""");
            Add(
                zip,
                "word/document.xml",
                $"""<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>""");
        }

        return buffer.ToArray();
    }

    static void Add(ZipArchive zip, string name, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open());
        writer.Write(content);
    }

    /// <summary>
    /// The text of a Word document, or null when the file is not one (yet), as while it is being written.
    /// </summary>
    public static string? Text(string path)
    {
        try
        {
            using var zip = ZipFile.OpenRead(path);
            using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
            var document = XDocument.Parse(reader.ReadToEnd());
            return string.Concat(document.Descendants().Where(_ => _.Name.LocalName == "t").Select(_ => _.Value));
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or System.Xml.XmlException or NullReferenceException)
        {
            return null;
        }
    }
}
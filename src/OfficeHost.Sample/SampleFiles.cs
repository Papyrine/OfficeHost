/// <summary>
/// The files the sample starts with.
/// </summary>
static class SampleFiles
{
    /// <summary>
    /// Adds a Word document to an empty folder, so that there is a file to open.
    /// </summary>
    public static void AddWelcome(string folder)
    {
        if (Directory.EnumerateFiles(folder).Any())
        {
            return;
        }

        using var file = File.Create(Path.Combine(folder, "Welcome.docx"));
        using var zip = new ZipArchive(file, ZipArchiveMode.Create);
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
            """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Welcome to the OfficeHost sample. Edit this document, and save it.</w:t></w:r></w:p></w:body></w:document>""");
    }

    static void Add(ZipArchive zip, string name, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open());
        writer.Write(content);
    }
}
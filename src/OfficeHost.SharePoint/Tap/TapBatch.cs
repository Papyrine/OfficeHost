using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

/// <summary>
/// The body of a batch the module sends the shadow (docs/shadow.md, "Protocol"): a <c>multipart/mixed</c> of a part per
/// item, gzip. An exchange's part is its request then its response, raw HTTP; a tombstone's part has no body.
/// </summary>
static class TapBatch
{
    public static byte[] Body(IReadOnlyList<TapItem> items, string boundary)
    {
        using var output = new MemoryStream();
        using (var gzip = new GZipStream(output, CompressionLevel.Fastest, true))
        {
            foreach (var item in items)
            {
                var head = new StringBuilder();
                head.Append("--").Append(boundary).Append("\r\n");
                Header(head, "X-Shadow-Sequence", item.Sequence.ToString(CultureInfo.InvariantCulture));
                Header(head, "X-Shadow-Started", item.Started.ToString("o", CultureInfo.InvariantCulture));
                if (item.IsTombstone)
                {
                    Header(head, "X-Shadow-Kind", "tombstone");
                    Header(head, "X-Shadow-Document", item.Document ?? "");
                    Header(head, "X-Shadow-Reason", item.Reason!);
                }
                else
                {
                    Header(head, "X-Shadow-Kind", "exchange");
                    Header(head, "X-Shadow-Request-Length", item.RequestLength.ToString(CultureInfo.InvariantCulture));
                }

                head.Append("\r\n");
                Write(gzip, head.ToString());
                if (!item.IsTombstone)
                {
                    gzip.Write(item.RequestHead, 0, item.RequestHead.Length);
                    item.RequestBody?.CopyTo(gzip);
                    gzip.Write(item.ResponseHead, 0, item.ResponseHead.Length);
                    item.ResponseBody?.CopyTo(gzip);
                }

                Write(gzip, "\r\n");
            }

            Write(gzip, $"--{boundary}--\r\n");
        }

        return output.ToArray();
    }

    static void Header(StringBuilder head, string name, string value) =>
        // header values of one line
        head.Append(name).Append(": ").Append(value.Replace("\r", " ").Replace("\n", " ")).Append("\r\n");

    static void Write(Stream stream, string text)
    {
        var bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}

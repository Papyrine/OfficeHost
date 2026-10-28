using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

/// <summary>
/// Records the exchanges to a folder, as the replay reads them (docs/replay.md): <c>{number}.request.http</c> and
/// <c>{number}.response.http</c>, numbered after those already there, the response first so that a request is never
/// read without it. Tombstones go to <c>tombstones.txt</c>. For a single server: two app domains writing the same folder
/// would number alike.
/// </summary>
sealed class FolderSink : ITapSink
{
    string folder;
    int number;

    public FolderSink(string folder)
    {
        this.folder = folder;
        Directory.CreateDirectory(folder);
        number = Directory.GetFiles(folder, "*.request.http").Length;
    }

    public TapSent Send(IReadOnlyList<TapItem> items, DateTime watermark, long dropped, bool final)
    {
        try
        {
            foreach (var item in items)
            {
                if (item.IsTombstone)
                {
                    File.AppendAllText(
                        Path.Combine(folder, "tombstones.txt"),
                        $"{item.Started.ToString("o", CultureInfo.InvariantCulture)} {item.Sequence} {item.Document} {item.Reason}\r\n");
                    continue;
                }

                var name = Interlocked.Increment(ref number).ToString("0000", CultureInfo.InvariantCulture);
                using (var response = File.Create(Path.Combine(folder, $"{name}.response.http")))
                {
                    response.Write(item.ResponseHead, 0, item.ResponseHead.Length);
                    item.ResponseBody?.CopyTo(response);
                }

                using var request = File.Create(Path.Combine(folder, $"{name}.request.http"));
                request.Write(item.RequestHead, 0, item.RequestHead.Length);
                item.RequestBody?.CopyTo(request);
            }

            if (dropped > 0)
            {
                File.AppendAllText(Path.Combine(folder, "tombstones.txt"), $"{TapClock.UtcNow.ToString("o", CultureInfo.InvariantCulture)} {dropped} dropped\r\n");
            }

            return new()
            {
                Sent = true
            };
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return new()
            {
                Error = exception.Message
            };
        }
    }
}

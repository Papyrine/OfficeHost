// The recorder of the SharePoint lab (docs/replay.md): compiled in the VM by Install-ReplayRecorder.ps1, with the C#
// compiler of the .NET Framework, so C# 5.
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Configuration;

[assembly: AssemblyVersion("1.0.0.0")]

/// <summary>
/// Records the cell storage and shared access requests of a SharePoint web application, and their responses, to a folder
/// (the appSetting ReplayRecorder.Folder), as the replay of OfficeHost reads them: <c>{number}.request.http</c> and
/// <c>{number}.response.http</c>, raw HTTP, numbered in the order the requests were authenticated. The request has
/// the user (<c>X-Replay-User</c>) and the time (<c>X-Replay-Time</c>), and no credentials. A request answered 401, a
/// step of the NTLM handshake, is not recorded. Errors of the recorder go to errors.txt in the folder, never to the
/// request.
/// </summary>
public class ReplayRecorder : IHttpModule
{
    const string defaultFolder = @"C:\ReplayRecording";

    static readonly string[] paths =
    {
        "/_vti_bin/cellstorage.svc/CellStorageService",
        "/_vti_bin/sharedaccess.asmx"
    };

    // the replay does not need them
    static readonly HashSet<string> secrets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "Authorization", "Proxy-Authorization", "Cookie", "Set-Cookie", "AuthMsaDeviceTicket", "X-IDCRL_AUTH_PARAMS_V1"
    };

    static readonly object gate = new object();
    static int number = -1;

    public void Init(HttpApplication application)
    {
        application.PostAuthenticateRequest += Start;
    }

    public void Dispose()
    {
    }

    static string Folder()
    {
        var folder = WebConfigurationManager.AppSettings["ReplayRecorder.Folder"];
        if (string.IsNullOrEmpty(folder))
        {
            return defaultFolder;
        }

        return folder;
    }

    static bool Recorded(HttpRequest request)
    {
        foreach (var path in paths)
        {
            if (request.Path.EndsWith(path, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The number of the next request: after those of the folder, the first time.
    /// </summary>
    static int Next(string folder)
    {
        lock (gate)
        {
            if (number < 0)
            {
                number = Directory.GetFiles(folder, "*.request.http").Length;
            }

            number++;
            return number;
        }
    }

    static void Start(object sender, EventArgs e)
    {
        var context = ((HttpApplication) sender).Context;
        if (!Recorded(context.Request))
        {
            return;
        }

        var folder = Folder();
        try
        {
            Directory.CreateDirectory(folder);
            var exchange = new Exchange(context, folder, Next(folder));
            context.Response.Filter = new Tee(context.Response.Filter, exchange);
        }
        catch (Exception exception)
        {
            Error(folder, exception);
        }
    }

    internal static void Error(string folder, Exception exception)
    {
        try
        {
            lock (gate)
            {
                File.AppendAllText(Path.Combine(folder, "errors.txt"), DateTime.UtcNow.ToString("o") + " " + exception + "\r\n");
            }
        }
        catch
        {
        }
    }

    /// <summary>
    /// A request, its headers taken when it was authenticated, written with its body and its response once the response
    /// is complete: the body is read after the handler read it, so that the recorder reads nothing before SharePoint.
    /// Reading it first, PowerPoint dropped the connection of its first request (docs/replay.md).
    /// </summary>
    class Exchange
    {
        readonly HttpContext context;
        readonly string folder;
        readonly int number;
        readonly byte[] requestHead;

        public Exchange(HttpContext context, string folder, int number)
        {
            this.context = context;
            this.folder = folder;
            this.number = number;
            var httpRequest = context.Request;
            var head = new StringBuilder();
            head.Append(httpRequest.HttpMethod).Append(' ').Append(httpRequest.RawUrl).Append(" HTTP/1.1\r\n");
            AppendHeaders(head, httpRequest.Headers);
            var user = "";
            if (context.User != null &&
                context.User.Identity != null)
            {
                user = context.User.Identity.Name;
            }

            head.Append("X-Replay-User: ").Append(user).Append("\r\n");
            head.Append("X-Replay-Time: ").Append(DateTime.UtcNow.ToString("o")).Append("\r\n\r\n");
            requestHead = Encoding.UTF8.GetBytes(head.ToString());
        }

        /// <summary>
        /// The body of the request, read again from its start: SharePoint's handler (cobalt.ashx for cellstorage.svc)
        /// read it through the buffered input stream.
        /// </summary>
        byte[] RequestBody()
        {
            try
            {
                var input = context.Request.InputStream;
                input.Position = 0;
                var body = new MemoryStream();
                input.CopyTo(body);
                return body.ToArray();
            }
            catch (Exception exception)
            {
                Error(folder, exception);
                return new byte[0];
            }
        }

        public void Write(byte[] body)
        {
            var response = context.Response;
            if (response.StatusCode == 401)
            {
                return;
            }

            var request = Concat(requestHead, RequestBody());

            var head = new StringBuilder();
            head.Append("HTTP/1.1 ").Append(response.StatusCode).Append(' ').Append(response.StatusDescription).Append("\r\n");
            var hasContentType = false;
            try
            {
                foreach (string name in response.Headers)
                {
                    if (name.Equals("Content-Type", StringComparison.OrdinalIgnoreCase))
                    {
                        hasContentType = true;
                    }
                }

                AppendHeaders(head, response.Headers);
            }
            catch (PlatformNotSupportedException)
            {
                // the headers of the response are there in the integrated pipeline only
            }

            if (!hasContentType &&
                response.ContentType != null)
            {
                var type = response.ContentType;
                if (response.Charset != null &&
                    response.Charset.Length > 0 &&
                    !type.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase))
                {
                    type += "; charset=" + response.Charset;
                }

                head.Append("Content-Type: ").Append(type).Append("\r\n");
            }

            head.Append("\r\n");
            var file = number.ToString("0000");
            // the response first: the replay reads a request with its response
            File.WriteAllBytes(Path.Combine(folder, file + ".response.http"), Concat(Encoding.UTF8.GetBytes(head.ToString()), body));
            File.WriteAllBytes(Path.Combine(folder, file + ".request.http"), request);
        }

        public string Folder
        {
            get { return folder; }
        }

        static void AppendHeaders(StringBuilder head, System.Collections.Specialized.NameValueCollection headers)
        {
            foreach (string name in headers)
            {
                if (name == null ||
                    secrets.Contains(name))
                {
                    continue;
                }

                head.Append(name).Append(": ").Append(string.Join(", ", headers.GetValues(name))).Append("\r\n");
            }
        }

        static byte[] Concat(byte[] first, byte[] second)
        {
            var result = new byte[first.Length + second.Length];
            Buffer.BlockCopy(first, 0, result, 0, first.Length);
            Buffer.BlockCopy(second, 0, result, first.Length, second.Length);
            return result;
        }
    }

    /// <summary>
    /// The response filter that keeps a copy of what the response sends, and writes the exchange when ASP.NET closes it,
    /// at the end of the response.
    /// </summary>
    class Tee : Stream
    {
        readonly Stream inner;
        readonly Exchange exchange;
        readonly MemoryStream copy = new MemoryStream();
        bool closed;

        public Tee(Stream inner, Exchange exchange)
        {
            this.inner = inner;
            this.exchange = exchange;
        }

        public override bool CanRead
        {
            get { return false; }
        }

        public override bool CanSeek
        {
            get { return false; }
        }

        public override bool CanWrite
        {
            get { return true; }
        }

        public override long Length
        {
            get { throw new NotSupportedException(); }
        }

        public override long Position
        {
            get { throw new NotSupportedException(); }
            set { throw new NotSupportedException(); }
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            copy.Write(buffer, offset, count);
            inner.Write(buffer, offset, count);
        }

        public override void Flush()
        {
            inner.Flush();
        }

        public override void Close()
        {
            if (!closed)
            {
                closed = true;
                try
                {
                    exchange.Write(copy.ToArray());
                }
                catch (Exception exception)
                {
                    Error(exchange.Folder, exception);
                }
            }

            inner.Close();
            base.Close();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }
    }
}

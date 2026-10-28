using System;
using System.Web;

namespace OfficeHost;

/// <summary>
/// The module of the OfficeHost shadow in SharePoint (docs/shadow.md): copies the cell storage and shared access requests
/// of desktop Office, and SharePoint's responses, to the shadow or a folder, after SharePoint has answered with them.
/// SharePoint's answer is never changed: the module reads the body of a cell storage request only after SharePoint's
/// handler has (that of IsOnlyClient, a few hundred bytes, before its ASMX handler, after which it reads empty), copies
/// what the response writes as it passes, sends from a thread of its own, and drops rather than wait or grow; its own
/// failures go to the event log, never to the request. Registered in web.config by Install-OfficeHostShadow.ps1.
/// </summary>
public sealed class ShadowModule : IHttpModule
{
    public void Init(HttpApplication application) =>
        application.PostAuthenticateRequest += Start;

    public void Dispose()
    {
    }

    static void Start(object? sender, EventArgs e)
    {
        try
        {
            var context = ((HttpApplication) sender!).Context;
            var path = context.Request.Path;
            if (TapCapture.KindOf(path) is not { } kind ||
                TapRuntime.Get() is not { } runtime ||
                !runtime.Copying ||
                !runtime.Settings.Includes(path))
            {
                return;
            }

            var capture = new TapCapture(context, runtime, kind);
            context.Response.Filter = new TapFilter(context.Response.Filter, capture);
            context.AddOnRequestCompleted(_ => capture.Completed());
        }
        catch (Exception exception)
        {
            TapLog.Write($"A request was not copied: {exception.Message}");
        }
    }
}

/// <summary>
/// Replays to the OfficeHost of the application in process: the handler of the desktop sites (<see cref="FssHttpHandler"/>,
/// as <see cref="OfficeHostExtensions.MapFssHttp"/> maps it), with a request of its own in a scope of its own, at the
/// time of the exchange (<see cref="ShadowClock"/>). No network, and none of the application's middleware
/// (authentication, HTTPS redirection) between the shadow and OfficeHost.
/// </summary>
sealed class InProcessTarget(IServiceScopeFactory scopes, IOptions<OfficeHostOptions> options) :
    IShadowTarget
{
    public async Task<HttpMessage> Send(string path, HttpMessage recorded, byte[] body, DateTimeOffset? time, Cancel cancel)
    {
        // {DesktopPath}/{token}/{rest}
        var desktopPath = options.Value.DesktopPath.TrimEnd('/');
        var relative = path[(desktopPath.Length + 1)..];
        var slash = relative.IndexOf('/');
        var token = relative[..slash];
        var rest = relative[(slash + 1)..];

        await using var scope = scopes.CreateAsyncScope();
        var context = new DefaultHttpContext
        {
            RequestServices = scope.ServiceProvider,
            RequestAborted = cancel
        };
        var request = context.Request;
        request.Method = HttpMethods.Post;
        request.Scheme = "https";
        request.Host = new("localhost");
        request.Path = path;
        request.ContentType = recorded.ContentType;
        request.ContentLength = body.Length;
        request.Body = new MemoryStream(body, false);
        if (recorded.Header("SOAPAction") is { } action)
        {
            request.Headers["SOAPAction"] = action;
        }

        using var responseBody = new MemoryStream();
        context.Response.Body = responseBody;
        var handler = scope.ServiceProvider.GetRequiredService<FssHttpHandler>();
        using (ShadowClock.At(time))
        {
            await handler.Handle(context, token, rest);
        }

        var response = context.Response;
        var headers = new List<(string Name, string Value)>();
        if (response.ContentType != null)
        {
            headers.Add(("Content-Type", response.ContentType));
        }

        return new($"HTTP/1.1 {response.StatusCode} {ReasonPhrases.GetReasonPhrase(response.StatusCode)}", headers, responseBody.ToArray());
    }
}

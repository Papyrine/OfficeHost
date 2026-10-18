/// <summary>
/// An HTTP response built up as the WOPI logic goes, the equivalent of a Flask <c>Response</c>.
/// </summary>
class WopiResult(int status) :
    IResult
{
    public int Status { get; set; } = status;

    public Dictionary<string, string> Headers { get; } = new(StringComparer.OrdinalIgnoreCase);

    public string ContentType { get; set; } = "text/html; charset=utf-8";

    public byte[] Body { get; set; } = [];

    public Stream? BodyStream { get; set; }

    /// <summary>
    /// A plain message, as Flask returns for a `(message, status)` tuple.
    /// </summary>
    public static WopiResult Text(string message, int status = StatusCodes.Status200OK) =>
        new(status)
        {
            Body = Encoding.UTF8.GetBytes(message)
        };

    public static WopiResult Json(string json, int status = StatusCodes.Status200OK) =>
        new(status)
        {
            ContentType = "application/json",
            Body = Encoding.UTF8.GetBytes(json)
        };

    public static WopiResult Json(JsonNode body, int status = StatusCodes.Status200OK) =>
        Json(body.ToJsonString(global::Json.Options), status);

    /// <summary>
    /// A JSON body with a single `message` property.
    /// </summary>
    public static WopiResult Message(string message, int status)
    {
        var result = Json(
            new JsonObject
            {
                ["message"] = message
            },
            status);
        if (status == StatusCodes.Status500InternalServerError)
        {
            result.Headers["X-WOPI-ServerError"] = message;
        }

        return result;
    }

    public async Task ExecuteAsync(HttpContext httpContext)
    {
        var response = httpContext.Response;
        response.StatusCode = Status;
        foreach (var (name, value) in Headers)
        {
            response.Headers[name] = value;
        }

        response.ContentType = ContentType;
        if (BodyStream != null)
        {
            await using (BodyStream)
            {
                await BodyStream.CopyToAsync(response.Body, httpContext.RequestAborted);
            }

            return;
        }

        response.ContentLength = Body.Length;
        await response.Body.WriteAsync(Body, httpContext.RequestAborted);
    }
}

static class HttpExtensions
{
    /// <summary>
    /// header used by reverse proxies such as traefik to pass the real remote IP address
    /// </summary>
    public const string RealIpHeader = "X-Real-IP";

    /// <summary>
    /// The first value of a query argument, or null when missing.
    /// </summary>
    public static string? Arg(this HttpRequest request, string name)
    {
        if (request.Query.TryGetValue(name, out var values) &&
            values.Count > 0)
        {
            return values[0];
        }

        return null;
    }

    /// <summary>
    /// The first value of a header, or null when missing.
    /// </summary>
    public static string? Header(this HttpRequest request, string name)
    {
        if (request.Headers.TryGetValue(name, out var values) &&
            values.Count > 0)
        {
            return values[0];
        }

        return null;
    }

    public static string RemoteAddress(this HttpContext context) =>
        context.Connection.RemoteIpAddress?.ToString() ?? "";

    /// <summary>
    /// The client address, as given by a reverse proxy when present.
    /// </summary>
    public static string Client(this HttpContext context) =>
        context.Request.Header(RealIpHeader) ?? context.RemoteAddress();

    /// <summary>
    /// The access token argument, shortened for logging.
    /// </summary>
    public static string TokenForLog(this HttpContext context) =>
        context.Request.Arg("access_token")?.Short() ?? "NA";

    /// <summary>
    /// The requested URL without the query, as Flask's <c>request.base_url</c>.
    /// </summary>
    public static string BaseUrl(this HttpRequest request) =>
        $"{request.Scheme}://{request.Host}{request.PathBase}{request.Path}";

    public static string Url(this HttpRequest request) =>
        $"{request.BaseUrl()}{request.QueryString}";
}

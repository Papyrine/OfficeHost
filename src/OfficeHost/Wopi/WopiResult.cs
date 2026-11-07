/// <summary>
/// A WOPI response: a status, headers, and a text, JSON or streamed body.
/// </summary>
class WopiResult(int status) :
    IResult
{
    public int Status { get; } = status;

    public Dictionary<string, string> Headers { get; } = new(StringComparer.OrdinalIgnoreCase);

    public string ContentType { get; init; } = "text/plain; charset=utf-8";

    public byte[] Body { get; init; } = [];

    public Stream? BodyStream { get; init; }

    public static WopiResult Ok() =>
        new(StatusCodes.Status200OK);

    public static WopiResult Text(string message, int status = StatusCodes.Status200OK) =>
        new(status)
        {
            Body = Encoding.UTF8.GetBytes(message)
        };

    public static WopiResult Json(JsonNode body, int status = StatusCodes.Status200OK) =>
        new(status)
        {
            ContentType = "application/json; charset=utf-8",
            Body = Encoding.UTF8.GetBytes(body.ToJsonString())
        };

    /// <summary>
    /// A server error, with the X-WOPI-ServerError header [MS-WOPI] 2.2.1 asks for.
    /// </summary>
    public static WopiResult ServerError(string message)
    {
        var result = Text(message, StatusCodes.Status500InternalServerError);
        result.Headers["X-WOPI-ServerError"] = message;
        return result;
    }

    public static WopiResult NotFound() =>
        Text("File not found", StatusCodes.Status404NotFound);

    public static WopiResult Unauthorized(string message) =>
        Text(message, StatusCodes.Status401Unauthorized);

    public static WopiResult NotImplemented(string message) =>
        Text(message, StatusCodes.Status501NotImplemented);

    /// <summary>
    /// A 409 for a lock mismatch: X-WOPI-Lock holds the current lock, empty when the file is not locked
    /// ([MS-WOPI] 3.3.5.1.13).
    /// </summary>
    public static WopiResult Conflict(string? currentLock, string reason)
    {
        var result = Text(reason, StatusCodes.Status409Conflict);
        result.Headers["X-WOPI-Lock"] = currentLock ?? "";
        result.Headers["X-WOPI-LockFailureReason"] = reason;
        return result;
    }

    public async Task ExecuteAsync(HttpContext httpContext)
    {
        // disposed even when a header is refused
        await using var bodyStream = BodyStream;
        var response = httpContext.Response;
        response.StatusCode = Status;
        foreach (var (name, value) in Headers)
        {
            response.Headers[name] = value;
        }

        if (Status == StatusCodes.Status500InternalServerError &&
            !response.Headers.ContainsKey("X-WOPI-ServerError"))
        {
            response.Headers["X-WOPI-ServerError"] = "Server error";
        }

        response.ContentType = ContentType;
        if (bodyStream != null)
        {
            // a length rather than a chunked response, when the stream knows it
            if (bodyStream.CanSeek)
            {
                response.ContentLength = bodyStream.Length - bodyStream.Position;
            }

            await bodyStream.CopyToAsync(response.Body, httpContext.RequestAborted);
            return;
        }

        response.ContentLength = Body.Length;
        await response.Body.WriteAsync(Body, httpContext.RequestAborted);
    }
}
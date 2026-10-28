/// <summary>
/// Records the requests of a client and their responses to a folder, as <see cref="Replayer"/> reads them, all as one
/// user, at the time of the application's clock: a recording of OfficeHost, to replay to another.
/// </summary>
class RecordingHandler(string folder, string user, TimeProvider time) : DelegatingHandler
{
    int number;

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, Cancel cancellationToken)
    {
        var requestBody = Array.Empty<byte>();
        if (request.Content != null)
        {
            requestBody = await request.Content.ReadAsByteArrayAsync(cancellationToken);
        }

        var requestHeaders = request.Headers
            .Concat(request.Content?.Headers ?? Enumerable.Empty<KeyValuePair<string, IEnumerable<string>>>())
            .Select(_ => (_.Key, string.Join(", ", _.Value)))
            .Append((Exchange.UserHeader, user))
            .Append((Exchange.TimeHeader, time.GetUtcNow().ToString("O")))
            .ToList();

        var response = await base.SendAsync(request, cancellationToken);
        var responseBody = await response.Content.ReadAsByteArrayAsync(cancellationToken);
        var responseHeaders = response.Content.Headers.Select(_ => (_.Key, string.Join(", ", _.Value))).ToList();

        var exchange = new Exchange(
            Interlocked.Increment(ref number),
            new($"{request.Method} {request.RequestUri!.PathAndQuery} HTTP/1.1", requestHeaders, requestBody),
            new($"HTTP/1.1 {(int) response.StatusCode} {response.ReasonPhrase}", responseHeaders, responseBody));
        exchange.Write(folder);

        // the body was read: the caller reads it again
        var content = new ByteArrayContent(responseBody);
        foreach (var (key, value) in response.Content.Headers)
        {
            content.Headers.TryAddWithoutValidation(key, value);
        }

        response.Content = content;
        return response;
    }
}

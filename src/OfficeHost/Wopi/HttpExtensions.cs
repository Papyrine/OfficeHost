static class HttpExtensions
{
    extension(HttpRequest request)
    {
        /// <summary>
        /// The first value of a header, or null when missing.
        /// </summary>
        public string? Header(string name)
        {
            if (request.Headers.TryGetValue(name, out var values) &&
                values.Count > 0)
            {
                return values[0];
            }

            return null;
        }

        /// <summary>
        /// The first value of a query parameter, or null when missing.
        /// </summary>
        public string? Query(string name)
        {
            if (request.Query.TryGetValue(name, out var values) &&
                values.Count > 0)
            {
                return values[0];
            }

            return null;
        }

        /// <summary>
        /// The URL the WOPI clients reach the application at: <see cref="OfficeHostOptions.BaseUrl"/>, or else the scheme, host
        /// and path base of the request.
        /// </summary>
        public string BaseUrl(OfficeHostOptions options)
        {
            if (options.BaseUrl != null)
            {
                return options.BaseUrl.TrimEnd('/');
            }

            return $"{request.Scheme}://{request.Host}{request.PathBase}";
        }

        public string DesktopBaseUrl(OfficeHostOptions options)
        {
            if (options.DesktopBaseUrl != null)
            {
                return options.DesktopBaseUrl.TrimEnd('/');
            }

            return $"{request.Scheme}://{request.Host}{request.PathBase}";
        }
    }
}
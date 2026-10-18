/// <summary>
/// HS256 JSON Web Tokens with PyJWT semantics. Hand-rolled because the Microsoft handlers reject HMAC keys
/// shorter than 256 bits, while the WOPI secrets are arbitrary strings shared with Reva and the WOPI proxy.
/// </summary>
static class Jwt
{
    static readonly string encodedHeader = Base64Url.EncodeToString("""{"alg":"HS256","typ":"JWT"}"""u8);

    public static string Encode<T>(T payload, string secret) =>
        Encode(JsonSerializer.SerializeToUtf8Bytes(payload, Json.Options), secret);

    public static string Encode(byte[] payload, string secret)
    {
        var signingInput = $"{encodedHeader}.{Base64Url.EncodeToString(payload)}";
        return $"{signingInput}.{Base64Url.EncodeToString(Sign(signingInput, secret))}";
    }

    static byte[] Sign(string signingInput, string secret) =>
        HMACSHA256.HashData(Encoding.UTF8.GetBytes(secret), Encoding.ASCII.GetBytes(signingInput));

    /// <summary>
    /// Validates the signature and, when present, the <c>exp</c> claim against <paramref name="now"/>
    /// (Unix seconds), then deserializes the payload.
    /// </summary>
    public static T Decode<T>(string token, string secret, long now)
    {
        var parts = token.Split('.');
        if (parts.Length != 3)
        {
            throw new JwtException("Not enough segments");
        }

        byte[] header;
        byte[] payload;
        byte[] signature;
        try
        {
            header = Base64Url.DecodeFromChars(parts[0]);
            payload = Base64Url.DecodeFromChars(parts[1]);
            signature = Base64Url.DecodeFromChars(parts[2]);
        }
        catch (FormatException)
        {
            throw new JwtException("Invalid crypto padding");
        }

        ValidateAlgorithm(header);
        var expected = Sign($"{parts[0]}.{parts[1]}", secret);
        if (!CryptographicOperations.FixedTimeEquals(expected, signature))
        {
            throw new JwtException("Signature verification failed");
        }

        try
        {
            ValidateExpiration(payload, now);
            return JsonSerializer.Deserialize<T>(payload, Json.Options) ??
                   throw new JwtException("Invalid payload");
        }
        catch (JsonException exception)
        {
            throw new JwtException($"Invalid payload: {exception.Message}");
        }
    }

    static void ValidateAlgorithm(byte[] header)
    {
        try
        {
            using var document = JsonDocument.Parse(header);
            if (document.RootElement.ValueKind == JsonValueKind.Object &&
                document.RootElement.TryGetProperty("alg", out var alg) &&
                alg.ValueKind == JsonValueKind.String &&
                alg.GetString() == "HS256")
            {
                return;
            }
        }
        catch (JsonException)
        {
            throw new JwtException("Invalid header string");
        }

        throw new JwtException("The specified alg value is not allowed");
    }

    static void ValidateExpiration(byte[] payload, long now)
    {
        using var document = JsonDocument.Parse(payload);
        if (document.RootElement.ValueKind != JsonValueKind.Object ||
            !document.RootElement.TryGetProperty("exp", out var exp))
        {
            return;
        }

        if (exp.ValueKind != JsonValueKind.Number ||
            !exp.TryGetInt64(out var seconds))
        {
            throw new JwtException("Expiration Time claim (exp) must be an integer.");
        }

        if (seconds <= now)
        {
            throw new JwtException("Signature has expired", expired: true);
        }
    }
}

class JwtException(string message, bool expired = false) :
    Exception(message)
{
    public bool Expired { get; } = expired;
}

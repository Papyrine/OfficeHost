/// <summary>
/// HS256 JSON Web Tokens with a required <c>exp</c> claim: all the built-in access tokens need, without a dependency.
/// </summary>
static class Jwt
{
    static string encodedHeader = Base64Url.EncodeToString("""{"alg":"HS256","typ":"JWT"}"""u8);

    static JsonSerializerOptions jsonOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static string Encode<T>(T payload, string key) =>
        Encode(JsonSerializer.SerializeToUtf8Bytes(payload, jsonOptions), key);

    public static string Encode(byte[] payload, string key)
    {
        var signingInput = $"{encodedHeader}.{Base64Url.EncodeToString(payload)}";
        return $"{signingInput}.{Base64Url.EncodeToString(Sign(signingInput, key))}";
    }

    static byte[] Sign(string signingInput, string key) =>
        HMACSHA256.HashData(Encoding.UTF8.GetBytes(key), Encoding.ASCII.GetBytes(signingInput));

    /// <summary>
    /// Validates the signature and the <c>exp</c> claim against <paramref name="now"/> (Unix seconds), then
    /// deserializes the payload.
    /// </summary>
    public static T Decode<T>(string token, string key, long now)
    {
        var parts = token.Split('.');
        if (parts.Length != 3)
        {
            throw new JwtException("Not a JSON Web Token");
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
            throw new JwtException("Invalid base64url");
        }

        ValidateAlgorithm(header);
        var expected = Sign($"{parts[0]}.{parts[1]}", key);
        if (!CryptographicOperations.FixedTimeEquals(expected, signature))
        {
            throw new JwtException("Signature verification failed");
        }

        try
        {
            ValidateExpiration(payload, now);
            return JsonSerializer.Deserialize<T>(payload, jsonOptions) ??
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
            throw new JwtException("Invalid header");
        }

        throw new JwtException("The alg is not HS256");
    }

    static void ValidateExpiration(byte[] payload, long now)
    {
        using var document = JsonDocument.Parse(payload);
        if (document.RootElement.ValueKind != JsonValueKind.Object ||
            !document.RootElement.TryGetProperty("exp", out var exp))
        {
            throw new JwtException("Missing exp");
        }

        if (exp.ValueKind != JsonValueKind.Number ||
            !exp.TryGetInt64(out var seconds))
        {
            throw new JwtException("The exp is not an integer");
        }

        if (seconds <= now)
        {
            throw new JwtException("Expired");
        }
    }
}

class JwtException(string message) :
    Exception(message);

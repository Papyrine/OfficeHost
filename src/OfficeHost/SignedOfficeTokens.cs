/// <summary>
/// Self-contained tokens: the access, signed as an HS256 JSON Web Token.
/// </summary>
sealed class SignedOfficeTokens(IOptions<OfficeHostOptions> options, TimeProvider time) :
    IOfficeTokens
{
    string key = ValidateKey(options.Value.SigningKey);
    TimeSpan lifetime = options.Value.TokenLifetime;

    static string ValidateKey(string? key)
    {
        if (key is {Length: >= 32})
        {
            return key;
        }

        throw new InvalidOperationException($"{nameof(OfficeHostOptions)}.{nameof(OfficeHostOptions.SigningKey)} must be at least 32 characters, or another {nameof(IOfficeTokens)} must be registered");
    }

    public Task<OfficeToken> Issue(OfficeAccess access, Cancel cancel = default)
    {
        var expires = DateTimeOffset.FromUnixTimeSeconds((time.GetUtcNow() + lifetime).ToUnixTimeSeconds());
        var claims = new Claims
        {
            FileId = access.FileId,
            UserId = access.UserId,
            UserName = access.UserName,
            IsAnonymous = access.IsAnonymous,
            CanWrite = access.CanWrite,
            CanWriteRelative = access.CanWriteRelative,
            CanRename = access.CanRename,
            ViewOnly = access.ViewOnly,
            Exp = expires.ToUnixTimeSeconds()
        };
        return Task.FromResult(new OfficeToken(Jwt.Encode(claims, key), access, expires));
    }

    public Task<OfficeToken?> Validate(string token, Cancel cancel = default)
    {
        Claims claims;
        try
        {
            claims = Jwt.Decode<Claims>(token, key, time.GetUtcNow().ToUnixTimeSeconds());
        }
        catch (JwtException)
        {
            return Task.FromResult<OfficeToken?>(null);
        }

        if (claims.FileId == null ||
            claims.UserId == null)
        {
            return Task.FromResult<OfficeToken?>(null);
        }

        var access = new OfficeAccess(claims.FileId, claims.UserId)
        {
            UserName = claims.UserName,
            IsAnonymous = claims.IsAnonymous,
            CanWrite = claims.CanWrite,
            CanWriteRelative = claims.CanWriteRelative,
            CanRename = claims.CanRename,
            ViewOnly = claims.ViewOnly
        };
        return Task.FromResult<OfficeToken?>(new(token, access, DateTimeOffset.FromUnixTimeSeconds(claims.Exp)));
    }

    sealed class Claims
    {
        [JsonPropertyName("fid")]
        public string? FileId { get; init; }

        [JsonPropertyName("uid")]
        public string? UserId { get; init; }

        [JsonPropertyName("name")]
        public string? UserName { get; init; }

        [JsonPropertyName("anon")]
        public bool IsAnonymous { get; init; }

        [JsonPropertyName("w")]
        public bool CanWrite { get; init; }

        [JsonPropertyName("wr")]
        public bool CanWriteRelative { get; init; }

        [JsonPropertyName("rn")]
        public bool CanRename { get; init; }

        [JsonPropertyName("vo")]
        public bool ViewOnly { get; init; }

        [JsonPropertyName("exp")]
        public long Exp { get; init; }
    }
}

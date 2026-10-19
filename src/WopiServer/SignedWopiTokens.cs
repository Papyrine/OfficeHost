/// <summary>
/// Self-contained tokens: the access, signed as an HS256 JSON Web Token.
/// </summary>
sealed class SignedWopiTokens(IOptions<WopiOptions> options, TimeProvider time) :
    IWopiTokens
{
    string key = ValidateKey(options.Value.SigningKey);
    TimeSpan lifetime = options.Value.TokenLifetime;

    static string ValidateKey(string? key)
    {
        if (key is {Length: >= 32})
        {
            return key;
        }

        throw new InvalidOperationException($"{nameof(WopiOptions)}.{nameof(WopiOptions.SigningKey)} must be at least 32 characters, or another {nameof(IWopiTokens)} must be registered");
    }

    public Task<WopiToken> Issue(WopiAccess access, Cancel cancel = default)
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
        return Task.FromResult(new WopiToken(Jwt.Encode(claims, key), access, expires));
    }

    public Task<WopiToken?> Validate(string token, Cancel cancel = default)
    {
        Claims claims;
        try
        {
            claims = Jwt.Decode<Claims>(token, key, time.GetUtcNow().ToUnixTimeSeconds());
        }
        catch (JwtException)
        {
            return Task.FromResult<WopiToken?>(null);
        }

        if (claims.FileId == null ||
            claims.UserId == null)
        {
            return Task.FromResult<WopiToken?>(null);
        }

        var access = new WopiAccess(claims.FileId, claims.UserId)
        {
            UserName = claims.UserName,
            IsAnonymous = claims.IsAnonymous,
            CanWrite = claims.CanWrite,
            CanWriteRelative = claims.CanWriteRelative,
            CanRename = claims.CanRename,
            ViewOnly = claims.ViewOnly
        };
        return Task.FromResult<WopiToken?>(new(token, access, DateTimeOffset.FromUnixTimeSeconds(claims.Exp)));
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

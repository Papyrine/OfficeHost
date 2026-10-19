/// <summary>
/// What an access token grants: a user's access to one file.
/// </summary>
/// <param name="FileId">The id of the file, as <see cref="IWopiStorage"/> knows it.</param>
/// <param name="UserId">A unique and stable id of the user.</param>
public record WopiAccess(string FileId, string UserId)
{
    /// <summary>
    /// The name of the user, as the WOPI client shows it.
    /// </summary>
    public string? UserName { get; init; }

    /// <summary>
    /// A user that is not signed in, such as the visitor of a public link.
    /// </summary>
    public bool IsAnonymous { get; init; }

    /// <summary>
    /// The user can change the file: lock it, save it, and delete it.
    /// </summary>
    public bool CanWrite { get; init; }

    /// <summary>
    /// The user can create files next to this one, as the Save As of the WOPI clients does (PutRelativeFile).
    /// </summary>
    public bool CanWriteRelative { get; init; }

    /// <summary>
    /// The user can rename the file.
    /// </summary>
    public bool CanRename { get; init; }

    /// <summary>
    /// The user can only view the file in the browser: no download, print or copy (RestrictedWebViewOnly).
    /// </summary>
    public bool ViewOnly { get; init; }
}

/// <summary>
/// An access token and what it grants.
/// </summary>
/// <param name="Value">The token, as the WOPI client passes it in the access_token parameter.</param>
public record WopiToken(string Value, WopiAccess Access, DateTimeOffset Expires);

/// <summary>
/// Issues and validates the access tokens. The built-in implementation signs the access in the token with
/// <see cref="WopiOptions.SigningKey"/>; register another one to use opaque tokens kept in a database, for instance.
/// </summary>
public interface IWopiTokens
{
    Task<WopiToken> Issue(WopiAccess access, Cancel cancel = default);

    /// <summary>
    /// The token, or null when it is invalid or expired.
    /// </summary>
    Task<WopiToken?> Validate(string token, Cancel cancel = default);
}

/// <summary>
/// Self-contained tokens: the access, signed as an HS256 JSON Web Token.
/// </summary>
sealed class SignedWopiTokens(IOptions<WopiOptions> options, TimeProvider time) :
    IWopiTokens
{
    readonly string key = ValidateKey(options.Value.SigningKey);
    readonly TimeSpan lifetime = options.Value.TokenLifetime;

    static string ValidateKey(string? key)
    {
        if (key == null ||
            key.Length < 32)
        {
            throw new InvalidOperationException($"{nameof(WopiOptions)}.{nameof(WopiOptions.SigningKey)} must be at least 32 characters, or another {nameof(IWopiTokens)} must be registered");
        }

        return key;
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

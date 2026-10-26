/// <summary>
/// Self-contained tokens: the access, signed with the signing key (<see cref="SignedToken"/>), and accepted when signed
/// with a previous one, so that the key can be rotated without ending the sessions.
/// </summary>
sealed class SignedOfficeTokens(IOptions<OfficeHostOptions> options, TimeProvider time) :
    IOfficeTokens
{
    byte[] key = Encoding.UTF8.GetBytes(ValidateKey(options.Value.SigningKey));
    byte[][] previousKeys = options.Value.PreviousSigningKeys.Select(_ => Encoding.UTF8.GetBytes(ValidateKey(_))).ToArray();
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
        return Task.FromResult(new OfficeToken(SignedToken.Encode(access, expires, key), access, expires));
    }

    public Task<OfficeToken?> Validate(string token, Cancel cancel = default)
    {
        var now = time.GetUtcNow();
        if (SignedToken.Decode(token, key, now) is var (access, expires))
        {
            return Task.FromResult<OfficeToken?>(new(token, access, expires));
        }

        foreach (var previous in previousKeys)
        {
            if (SignedToken.Decode(token, previous, now) is var (previousAccess, previousExpires))
            {
                return Task.FromResult<OfficeToken?>(new(token, previousAccess, previousExpires));
            }
        }

        return Task.FromResult<OfficeToken?>(null);
    }
}

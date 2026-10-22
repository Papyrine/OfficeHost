/// <summary>
/// Self-contained tokens: the access, signed with the signing key (<see cref="SignedToken"/>).
/// </summary>
sealed class SignedOfficeTokens(IOptions<OfficeHostOptions> options, TimeProvider time) :
    IOfficeTokens
{
    byte[] key = Encoding.UTF8.GetBytes(ValidateKey(options.Value.SigningKey));
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
        if (SignedToken.Decode(token, key, time.GetUtcNow()) is not var (access, expires))
        {
            return Task.FromResult<OfficeToken?>(null);
        }

        return Task.FromResult<OfficeToken?>(new(token, access, expires));
    }
}

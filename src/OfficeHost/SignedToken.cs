/// <summary>
/// The format of the built-in access tokens: the access and when it expires, in bytes, then the first 16 bytes of their
/// HMAC-SHA256 with the signing key, all in base64url. Short, so that the URL of a desktop site, which holds the token,
/// stays under the 255 characters Word accepts for a Save As (docs/fsshttp-notes.md):
/// - a version, 1;
/// - flags: anonymous (0x01), write (0x02), write relative (0x04), rename (0x08), view only (0x10), with a user name
///   (0x20);
/// - the expiry, in Unix seconds, 4 bytes, little-endian;
/// - the file id, the user id and the user name when there is one: each its length in UTF-8 bytes as a 7-bit encoded
///   integer, then the bytes.
/// </summary>
static class SignedToken
{
    const byte version = 1;
    const int macLength = 16;
    // a token longer than this is not one: no work is done on it
    const int maxLength = 4096;

    const byte anonymous = 0x01;
    const byte write = 0x02;
    const byte writeRelative = 0x04;
    const byte rename = 0x08;
    const byte viewOnly = 0x10;
    const byte userName = 0x20;

    public static string Encode(OfficeAccess access, DateTimeOffset expires, byte[] key)
    {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, true))
        {
            writer.Write(version);
            writer.Write(Flags(access));
            writer.Write(checked((uint) expires.ToUnixTimeSeconds()));
            writer.Write(access.FileId);
            writer.Write(access.UserId);
            if (access.UserName != null)
            {
                writer.Write(access.UserName);
            }
        }

        var payload = stream.ToArray();
        return Base64Url.EncodeToString([.. payload, .. Mac(payload, key)]);
    }

    static byte Flags(OfficeAccess access)
    {
        byte flags = 0;
        if (access.IsAnonymous)
        {
            flags |= anonymous;
        }

        if (access.CanWrite)
        {
            flags |= write;
        }

        if (access.CanWriteRelative)
        {
            flags |= writeRelative;
        }

        if (access.CanRename)
        {
            flags |= rename;
        }

        if (access.ViewOnly)
        {
            flags |= viewOnly;
        }

        if (access.UserName != null)
        {
            flags |= userName;
        }

        return flags;
    }

    /// <summary>
    /// The access of the token and when it expires, or null when it is not a token signed with the key, or has expired
    /// at <paramref name="now"/>.
    /// </summary>
    public static (OfficeAccess Access, DateTimeOffset Expires)? Decode(string token, byte[] key, DateTimeOffset now)
    {
        if (token.Length > maxLength)
        {
            return null;
        }

        byte[] bytes;
        try
        {
            bytes = Base64Url.DecodeFromChars(token);
        }
        catch (FormatException)
        {
            return null;
        }

        if (bytes.Length <= macLength)
        {
            return null;
        }

        var payload = bytes.AsSpan(0, bytes.Length - macLength);
        if (!CryptographicOperations.FixedTimeEquals(Mac(payload, key), bytes.AsSpan(payload.Length)))
        {
            return null;
        }

        try
        {
            using var reader = new BinaryReader(new MemoryStream(payload.ToArray()), Encoding.UTF8);
            if (reader.ReadByte() != version)
            {
                return null;
            }

            var flags = reader.ReadByte();
            var expires = DateTimeOffset.FromUnixTimeSeconds(reader.ReadUInt32());
            var fileId = reader.ReadString();
            var userId = reader.ReadString();
            string? name = null;
            if ((flags & userName) != 0)
            {
                name = reader.ReadString();
            }

            if (reader.BaseStream.Position != reader.BaseStream.Length ||
                expires <= now)
            {
                return null;
            }

            var access = new OfficeAccess(fileId, userId)
            {
                UserName = name,
                IsAnonymous = (flags & anonymous) != 0,
                CanWrite = (flags & write) != 0,
                CanWriteRelative = (flags & writeRelative) != 0,
                CanRename = (flags & rename) != 0,
                ViewOnly = (flags & viewOnly) != 0
            };
            return (access, expires);
        }
        catch (EndOfStreamException)
        {
            return null;
        }
    }

    static byte[] Mac(ReadOnlySpan<byte> payload, byte[] key) =>
        HMACSHA256.HashData(key, payload)[..macLength];
}

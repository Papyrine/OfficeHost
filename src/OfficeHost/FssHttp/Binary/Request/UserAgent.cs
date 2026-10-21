/// <summary>
/// The user agent of a request ([MS-FSSHTTPB] 2.2.2): a GUID, a client and platform (which replace the GUID, though
/// Word sends both), and a version.
/// </summary>
record UserAgent(Guid? Guid, string? Client, string? Platform, uint Version)
{
    public static UserAgent Read(ref FssHttpBReader reader)
    {
        reader.ReadStart(StreamObjectType.UserAgent).EnsureEnd("a user agent start");
        Guid? guid = null;
        if (reader.PeekStart(StreamObjectType.UserAgentGuid))
        {
            var fields = reader.ReadStart(StreamObjectType.UserAgentGuid);
            guid = fields.ReadGuid();
            fields.EnsureEnd("a user agent GUID");
        }

        string? client = null;
        string? platform = null;
        if (reader.PeekStart(StreamObjectType.UserAgentClientAndPlatform))
        {
            var fields = reader.ReadStart(StreamObjectType.UserAgentClientAndPlatform);
            client = Encoding.UTF8.GetString(fields.ReadBinaryItem().Span);
            platform = Encoding.UTF8.GetString(fields.ReadBinaryItem().Span);
            fields.EnsureEnd("a user agent client and platform");
        }

        var version = reader.ReadStart(StreamObjectType.UserAgentVersion);
        var value = version.ReadUInt32();
        version.EnsureEnd("a user agent version");
        reader.ReadEnd(StreamObjectType.UserAgent);
        return new(guid, client, platform, value);
    }

    public void Write(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.UserAgent, 0);
        if (Guid != null)
        {
            writer.WriteStart(StreamObjectType.UserAgentGuid, 16);
            writer.WriteGuid(Guid.Value);
        }

        if (Client != null || Platform != null)
        {
            var client = Encoding.UTF8.GetBytes(Client ?? "");
            var platform = Encoding.UTF8.GetBytes(Platform ?? "");
            writer.WriteStart(StreamObjectType.UserAgentClientAndPlatform, FssHttpBSize.BinaryItem(client.Length) + FssHttpBSize.BinaryItem(platform.Length));
            writer.WriteBinaryItem(client);
            writer.WriteBinaryItem(platform);
        }

        writer.WriteStart(StreamObjectType.UserAgentVersion, 4);
        writer.WriteUInt32(Version);
        writer.WriteEnd(StreamObjectType.UserAgent);
    }
}
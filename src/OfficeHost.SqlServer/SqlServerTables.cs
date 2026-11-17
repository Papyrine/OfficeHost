/// <summary>
/// The tables of the SQL Server stores, shared with OfficeHost.SqlServer.Install, which creates them.
/// </summary>
static class SqlServerTables
{
    public const string Files = "OfficeHostFiles";
    public const string Locks = "OfficeHostLocks";
    public const string FssHttpStates = "OfficeHostFssHttpStates";
    public const string UserInfo = "OfficeHostUserInfo";
    public const string SchemaVersion = "OfficeHostSchemaVersion";

    /// <summary>
    /// The longest file id of a coauthoring state: with its Partition, 16 bytes, its key fits the 900 bytes of a
    /// clustered index.
    /// </summary>
    public const int MaxStateFileIdLength = 442;

    public static string Quote(string identifier) =>
        $"[{identifier.Replace("]", "]]")}]";

    public static string Name(string schema, string table) =>
        $"{Quote(schema)}.{Quote(table)}";
}

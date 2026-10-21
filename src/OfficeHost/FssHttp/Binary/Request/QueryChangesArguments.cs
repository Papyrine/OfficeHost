/// <summary>
/// The Query Changes Request Arguments ([MS-FSSHTTPB] 2.2.2.1.3): a byte of flags (F, include the Storage Manifest,
/// 0x01; G, include cell changes, 0x02), and the cell the query is scoped to, the null Cell ID for none.
/// </summary>
record QueryChangesArguments(byte Flags, CellId? Cell)
{
    public bool IncludeStorageManifest => (Flags & 0x01) != 0;

    public bool IncludeCellChanges => (Flags & 0x02) != 0;
}
/// <summary>
/// How the signatures of zip items are made ([MS-FSSHTTPD] 2.4.1).
/// </summary>
/// <param name="XorItemSignatures">The signature of an item of at most 4096 bytes, header and data, is the exclusive or of
/// the signatures of its header and its data, as between clients and servers of version 2.2 or greater ([MS-FSSHTTP]
/// 2.2.5.10, 2.2.5.13), not their concatenation.</param>
/// <param name="CompressedDataCrc">The signature of the data of an item has the CRC-32 of its compressed data after the
/// CRC of the item, as the note &lt;6&gt; says Office does from the Office 2010 March 2015 CU. Word 16.0.20326 does not
/// (docs/fsshttp-notes.md).</param>
record ChunkingOptions(bool XorItemSignatures = true, bool CompressedDataCrc = false)
{
    /// <summary>
    /// The signatures of Word 16.0.20326 saving to SharePoint Server, and of this host.
    /// </summary>
    public static ChunkingOptions Default { get; } = new();

    /// <summary>
    /// The signatures of the example of [MS-FSSHTTPD] 3.1: concatenated, without the CRC of the compressed data.
    /// </summary>
    public static ChunkingOptions Original { get; } = new(false);
}
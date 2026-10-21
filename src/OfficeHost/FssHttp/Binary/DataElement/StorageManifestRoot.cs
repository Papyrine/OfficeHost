/// <summary>
/// A root of a Storage Manifest: a root identifier and its cell.
/// </summary>
record StorageManifestRoot(ExtendedGuid RootId, CellId Cell);
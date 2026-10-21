/// <summary>
/// A Revision Manifest data element ([MS-FSSHTTPB] 2.2.1.12.5): a revision, the revision it is based on, its roots and
/// its Object Groups.
/// </summary>
record RevisionManifest(
    ExtendedGuid Id,
    SerialNumber SerialNumber,
    ExtendedGuid RevisionId,
    ExtendedGuid BaseRevisionId,
    IReadOnlyList<RevisionManifestRoot> Roots,
    IReadOnlyList<ExtendedGuid> ObjectGroups) :
    DataElement(Id, SerialNumber)
{
    public override DataElementType Type => DataElementType.RevisionManifest;

    public static RevisionManifest ReadBody(ref FssHttpBReader reader, ExtendedGuid id, SerialNumber serialNumber)
    {
        var fields = reader.ReadStart(StreamObjectType.RevisionManifest);
        var revisionId = fields.ReadExtendedGuid();
        var baseRevisionId = fields.ReadExtendedGuid();
        fields.EnsureEnd("a Revision Manifest");
        var roots = new List<RevisionManifestRoot>();
        while (reader.PeekStart(StreamObjectType.RevisionManifestRootDeclare))
        {
            var root = reader.ReadStart(StreamObjectType.RevisionManifestRootDeclare);
            roots.Add(new(root.ReadExtendedGuid(), root.ReadExtendedGuid()));
            root.EnsureEnd("a Revision Manifest root declare");
        }

        var groups = new List<ExtendedGuid>();
        while (reader.PeekStart(StreamObjectType.RevisionManifestObjectGroupReferences))
        {
            var group = reader.ReadStart(StreamObjectType.RevisionManifestObjectGroupReferences);
            groups.Add(group.ReadExtendedGuid());
            group.EnsureEnd("a Revision Manifest Object Group reference");
        }

        return new(id, serialNumber, revisionId, baseRevisionId, roots, groups);
    }

    protected override void WriteBody(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.RevisionManifest, RevisionId.Size + BaseRevisionId.Size);
        writer.WriteExtendedGuid(RevisionId);
        writer.WriteExtendedGuid(BaseRevisionId);
        foreach (var root in Roots)
        {
            writer.WriteStart(StreamObjectType.RevisionManifestRootDeclare, root.RootId.Size + root.ObjectId.Size);
            writer.WriteExtendedGuid(root.RootId);
            writer.WriteExtendedGuid(root.ObjectId);
        }

        foreach (var group in ObjectGroups)
        {
            writer.WriteStart(StreamObjectType.RevisionManifestObjectGroupReferences, group.Size);
            writer.WriteExtendedGuid(group);
        }
    }
}
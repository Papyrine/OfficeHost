record ObjectDataBlobReference(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences, ExtendedGuid BlobId) :
    ObjectGroupData(ObjectReferences, CellReferences);
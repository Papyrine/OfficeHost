record ObjectData(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences, ReadOnlyMemory<byte> Data) :
    ObjectGroupData(ObjectReferences, CellReferences);
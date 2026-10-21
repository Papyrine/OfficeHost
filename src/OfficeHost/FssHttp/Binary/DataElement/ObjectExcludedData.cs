record ObjectExcludedData(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences, ulong DataSize) :
    ObjectGroupData(ObjectReferences, CellReferences);
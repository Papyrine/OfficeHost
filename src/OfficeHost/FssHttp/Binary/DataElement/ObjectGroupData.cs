/// <summary>
/// The data of an object of an Object Group: Object Data (2.2.1.12.6.4) with its data, Object Excluded Data with the
/// size of the data only, or an Object Data BLOB Reference (2.2.1.12.6.5).
/// </summary>
abstract record ObjectGroupData(IReadOnlyList<ExtendedGuid> ObjectReferences, IReadOnlyList<CellId> CellReferences);
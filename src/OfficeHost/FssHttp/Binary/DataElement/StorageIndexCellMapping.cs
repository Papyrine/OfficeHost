/// <summary>
/// A Storage Index Cell Mapping: the Cell Manifest of a cell.
/// </summary>
record StorageIndexCellMapping(CellId Cell, ExtendedGuid MappingId, SerialNumber MappingSerialNumber) :
    StorageIndexEntry(MappingId, MappingSerialNumber);
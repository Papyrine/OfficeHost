/// <summary>
/// The data element types of [MS-FSSHTTPB] 2.2.1.12.1.
/// </summary>
enum DataElementType : ulong
{
    StorageIndex = 1,
    StorageManifest = 2,
    CellManifest = 3,
    RevisionManifest = 4,
    ObjectGroup = 5,
    DataElementFragment = 6,
    ObjectDataBlob = 10
}
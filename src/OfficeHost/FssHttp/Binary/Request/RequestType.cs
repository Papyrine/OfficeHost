/// <summary>
/// The request types of [MS-FSSHTTPB] 2.2.1.6.
/// </summary>
enum RequestType : ulong
{
    QueryAccess = 1,
    QueryChanges = 2,
    PutChanges = 5,
    AllocateExtendedGuidRange = 11
}
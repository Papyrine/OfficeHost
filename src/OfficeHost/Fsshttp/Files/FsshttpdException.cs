/// <summary>
/// Data elements that do not hold a file stream as [MS-FSSHTTPD] represents it: a missing or unexpected data element or
/// object, a tree of node objects that is not one, or sizes that do not add up.
/// </summary>
class FsshttpdException(string message, Exception? inner = null) :
    Exception(message, inner);

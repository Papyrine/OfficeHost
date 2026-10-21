/// <summary>
/// A range of Serial Numbers: the GUID with the values from From to To.
/// </summary>
record CellKnowledgeRange(Guid Guid, ulong From, ulong To) :
    CellKnowledgeItem;
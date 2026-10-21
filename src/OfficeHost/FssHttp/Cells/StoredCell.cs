/// <summary>
/// A cell state as stored, with the version of the <see cref="FssHttpState"/> it came from, null when there was none.
/// </summary>
record StoredCell(string? Version, CellState State);
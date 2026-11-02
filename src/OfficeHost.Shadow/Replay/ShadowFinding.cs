/// <summary>
/// What the shadow found of a sub-response, or of a whole exchange (a label of <c>cell storage</c> or
/// <c>IsOnlyClient</c>).
/// </summary>
/// <param name="Detail">The error code of a sub-response the same on both.</param>
/// <param name="Reason">Why it is a known difference or was not compared.</param>
/// <param name="ContentCompared">Whether the files of both copies of a partition were compared after it.</param>
/// <param name="Files">The two files of a content query whose files differ.</param>
record ShadowFinding(
    Exchange Exchange,
    string? Document,
    string Label,
    ShadowOutcome Outcome,
    IReadOnlyList<ShadowDifference> Differences,
    string? Detail = null,
    string? Reason = null,
    bool ContentCompared = false,
    (byte[] SharePoint, byte[] OfficeHost)? Files = null);
namespace OfficeHost;

/// <summary>
/// A difference between SharePoint and OfficeHost that is expected: a sub-response whose differences all match rules is
/// a <see cref="ShadowOutcome.KnownDifference"/>, not <see cref="ShadowOutcome.Different"/>. Each pattern is a regular
/// expression; a null one matches anything.
/// </summary>
/// <param name="Label">The sub-request, as <c>GetVersions</c>, <c>Coauth JoinCoauthoring</c> or
/// <c>Cell QueryChanges (translated)</c>.</param>
/// <param name="Facet">What differs: <c>ErrorCode</c>, <c>ErrorMessage</c>, an attribute (<c>@Etag</c>), a list of
/// properties (<c>DocProps</c>), <c>FSSHTTPB</c>, <c>content</c>, <c>editors</c>.</param>
/// <param name="Reason">Why it is expected, for the log.</param>
/// <param name="SharePoint">SharePoint's value, or (none).</param>
/// <param name="OfficeHost">OfficeHost's value, or (none).</param>
public record ShadowRule(string? Label, string? Facet, string Reason, string? SharePoint = null, string? OfficeHost = null)
{
    /// <summary>
    /// The differences the lab found expected (docs/fsshttp-notes.md, "Word's session replayed").
    /// </summary>
    public static IReadOnlyList<ShadowRule> Defaults { get; } =
    [
        new("^GetVersions$", "^ErrorCode$", "OfficeHost does not support GetVersions", "^Success$", "^RequestNotSupported$"),
        new("^GetDocMetaInfo$", "^(DocProps|FolderProps)$", "OfficeHost sends a subset of SharePoint's vti_ properties"),
        new("^Cell QueryChanges", "^@HaveOnlyDemotionChanges$", "SharePoint sends a revision of its own after a save", "^False$", @"^\(none\)$")
    ];

    public bool Matches(string label, string facet, string? sharePoint, string? officeHost) =>
        IsMatch(Label, label) &&
        IsMatch(Facet, facet) &&
        IsMatch(SharePoint, sharePoint ?? "(none)") &&
        IsMatch(OfficeHost, officeHost ?? "(none)");

    static bool IsMatch(string? pattern, string value) =>
        pattern == null ||
        Regex.IsMatch(value, pattern, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
}

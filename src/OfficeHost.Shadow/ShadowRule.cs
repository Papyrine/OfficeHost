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
    /// The differences found expected: in the lab, against SharePoint Server SE (docs/fsshttp-notes.md, "Word's session
    /// replayed"), and where SharePoint Server 2016 answers otherwise than SE, which OfficeHost follows
    /// (docs/fsshttp-notes.md, "SharePoint Server 2016").
    /// </summary>
    public static IReadOnlyList<ShadowRule> Defaults { get; } =
    [
        new("^GetVersions$", "^ErrorCode$", "OfficeHost does not support GetVersions", "^Success$", "^RequestNotSupported$"),
        new("^GetDocMetaInfo$", "^(DocProps|FolderProps)$", "OfficeHost sends a subset of SharePoint's vti_ properties"),
        new(
            "^Cell (QueryChanges|QueryAccess|PutChanges)",
            "^@HaveOnlyDemotionChanges$",
            "SharePoint sends it where OfficeHost does not: SE with a revision of its own after a save, 2016 with every Cell of the content",
            "^False$",
            @"^\(none\)$"),
        new(
            "^Coauth MarkTransitionComplete$",
            "^ErrorCode$",
            "SharePoint 2016 answers InvalidCoauthSession for a client out of the session, SE and OfficeHost Success",
            "^InvalidCoauthSession$",
            "^Success$"),
        new("^Cell ", "^editors JoinCoauthTime$", "SharePoint 2016 adds no JoinCoauthTime to the editors table, SE and OfficeHost do", "^$"),
        new(
            "^Cell ",
            "^editors HasEditorPermission$",
            "SharePoint 2016 drops HasEditorPermission from the editors table when a client updates its metadata",
            @"^\{[0-9A-F-]+\}=(; \{[0-9A-F-]+\}=)*$",
            @"^\{[0-9A-F-]+\}=true(; \{[0-9A-F-]+\}=true)*$"),
        new("^Cell ", "^editors$", "SharePoint 2016 sends no editors table when no one is editing, OfficeHost an empty one", @"^\(none\)$", "^$")
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

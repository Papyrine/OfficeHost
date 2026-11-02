/// <summary>
/// A difference of a SubResponse between SharePoint and OfficeHost: the facet (<see cref="ShadowComparer.Facets"/>) and
/// each server's value of it, or null when it had none.
/// </summary>
record ShadowDifference(string Facet, string? SharePoint, string? OfficeHost)
{
    public override string ToString()
    {
        if (SharePoint != null &&
            OfficeHost != null &&
            !Facet.StartsWith('@') &&
            Facet is not ("ErrorCode" or "ErrorMessage" or "FSSHTTPB" or "file" or "content" or "editors" or "HTTP" or "IsOnlyClientResult"))
        {
            // a list of keys: what one has and the other not
            var sharePoint = SharePoint.Split(',');
            var officeHost = OfficeHost.Split(',');
            return $"{Facet}: only SharePoint {string.Join(',', sharePoint.Except(officeHost))}; only OfficeHost {string.Join(",", officeHost.Except(sharePoint))}";
        }

        return $"{Facet}: {SharePoint ?? "(none)"} / {OfficeHost ?? "(none)"}";
    }
}
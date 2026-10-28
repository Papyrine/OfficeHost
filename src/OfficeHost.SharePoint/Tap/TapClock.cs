/// <summary>
/// The time of the module: UTC from a stopwatch anchored once, so that it never goes back within an app domain, as the
/// watermarks it sends need (docs/shadow.md, "Order").
/// </summary>
static class TapClock
{
    static DateTime anchor = DateTime.UtcNow;
    static Stopwatch stopwatch = Stopwatch.StartNew();

    public static DateTime UtcNow => anchor + stopwatch.Elapsed;
}

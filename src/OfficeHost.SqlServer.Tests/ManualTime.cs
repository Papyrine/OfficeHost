/// <summary>
/// A clock the tests move by hand, so lock expirations need no sleeping.
/// </summary>
class ManualTime : TimeProvider
{
    public DateTimeOffset Now { get; set; } = DateTimeOffset.UtcNow;

    public override DateTimeOffset GetUtcNow() =>
        Now;

    public void Advance(TimeSpan span) =>
        Now += span;
}

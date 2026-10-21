static class SqlExceptionExtensions
{
    /// <summary>
    /// A violation of a primary key or of a unique index.
    /// </summary>
    public static bool IsDuplicateKey(this SqlException exception) =>
        exception.Number is 2627 or 2601;
}

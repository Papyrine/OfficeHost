/// <summary>
/// The log lines of an application, as `Level Category: message`.
/// </summary>
sealed class TestLog : ILoggerProvider
{
    List<string> lines = [];

    public IReadOnlyList<string> Lines
    {
        get
        {
            lock (lines)
            {
                return lines.ToList();
            }
        }
    }

    public override string ToString() =>
        string.Join('\n', Lines);

    public ILogger CreateLogger(string categoryName) =>
        new Logger(this, categoryName);

    public void Dispose()
    {
    }

    sealed class Logger(TestLog log, string category) : ILogger
    {
        public IDisposable? BeginScope<TState>(TState state)
            where TState : notnull =>
            null;

        public bool IsEnabled(LogLevel logLevel) =>
            true;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
        {
            lock (log.lines)
            {
                log.lines.Add($"{logLevel} {category}: {formatter(state, exception)}");
            }
        }
    }
}
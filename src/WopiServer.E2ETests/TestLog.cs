/// <summary>
/// The log lines of the application.
/// </summary>
sealed class TestLog : ILoggerProvider
{
    List<string> lines = [];

    public override string ToString()
    {
        lock (lines)
        {
            return string.Join('\n', lines);
        }
    }

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
                log.lines.Add($"{logLevel} {category}: {formatter(state, exception)}{(exception == null ? "" : $" {exception}")}");
            }
        }
    }
}
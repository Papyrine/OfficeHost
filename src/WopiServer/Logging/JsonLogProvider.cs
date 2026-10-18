/// <summary>
/// Writes one JSON object per log entry, in the format of the Python implementation:
/// <c>{"time": "...", "host": "...", "level": "INFO", "process": "wopiserver", "module": "wopi", "msg": "...", "key": "value", ...}</c>.
/// The message is the literal text of the template up to its first placeholder, and each placeholder becomes a key.
/// </summary>
sealed class JsonLogProvider(TextWriter writer, bool ownsWriter, string host, Func<LogLevel> minimumLevel) :
    ILoggerProvider
{
    readonly Lock sync = new();

    public ILogger CreateLogger(string categoryName) =>
        new JsonLogger(this, categoryName);

    public void Dispose()
    {
        lock (sync)
        {
            writer.Flush();
            if (ownsWriter)
            {
                writer.Dispose();
            }
        }
    }

    /// <summary>
    /// The configured log level applies to the WOPI server, framework categories only log warnings and above.
    /// </summary>
    public bool IsEnabled(string category, LogLevel level)
    {
        if (level == LogLevel.None)
        {
            return false;
        }

        var minimum = minimumLevel();
        if (category.StartsWith("Microsoft.", StringComparison.Ordinal) ||
            category.StartsWith("System.", StringComparison.Ordinal))
        {
            minimum = (LogLevel) Math.Max((int) minimum, (int) LogLevel.Warning);
        }

        return level >= minimum;
    }

    static string LevelName(LogLevel level) =>
        level switch
        {
            LogLevel.Trace or LogLevel.Debug => "DEBUG",
            LogLevel.Information => "INFO",
            LogLevel.Warning => "WARNING",
            LogLevel.Error => "ERROR",
            _ => "CRITICAL"
        };

    public void Write(string category, LogLevel level, string message, IReadOnlyList<KeyValuePair<string, object?>>? values, Exception? exception)
    {
        using var buffer = new MemoryStream();
        using (var json = new Utf8JsonWriter(buffer))
        {
            json.WriteStartObject();
            json.WriteString("time", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture));
            json.WriteString("host", host);
            json.WriteString("level", LevelName(level));
            json.WriteString("process", "wopiserver");
            json.WriteString("module", category);
            if (values?.FirstOrDefault(_ => _.Key == "{OriginalFormat}").Value is not string format)
            {
                json.WriteString("msg", message);
            }
            else
            {
                json.WriteString("msg", Message(format, values!));
                foreach (var (key, value) in values!)
                {
                    if (key == "{OriginalFormat}" ||
                        key == "msg")
                    {
                        continue;
                    }

                    json.WriteString(key, Format(value));
                }
            }

            if (exception != null)
            {
                json.WriteString("exception", exception.ToString());
            }

            json.WriteEndObject();
        }

        var line = Encoding.UTF8.GetString(buffer.ToArray());
        lock (sync)
        {
            writer.WriteLine(line);
            writer.Flush();
        }
    }

    /// <summary>
    /// The literal text of the template up to its first placeholder. Placeholders leading the template are rendered,
    /// so `{operation}: client context {trace}` gives `Checkfileinfo: client context`.
    /// </summary>
    static string Message(string format, IReadOnlyList<KeyValuePair<string, object?>> values)
    {
        var builder = new StringBuilder();
        var hasText = false;
        var index = 0;
        while (index < format.Length)
        {
            var c = format[index];
            if (c == '{')
            {
                var end = format.IndexOf('}', index);
                if (hasText ||
                    end < 0)
                {
                    break;
                }

                var name = format[(index + 1)..end].Split(':', ',')[0];
                builder.Append(Format(values.FirstOrDefault(_ => _.Key == name).Value));
                index = end + 1;
                continue;
            }

            builder.Append(c);
            hasText |= !char.IsWhiteSpace(c);
            index++;
        }

        return builder.ToString().Trim();
    }

    static string Format(object? value) =>
        value switch
        {
            null => "None",
            bool flag => flag ? "True" : "False",
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString() ?? ""
        };

    sealed class JsonLogger(JsonLogProvider provider, string category) :
        ILogger
    {
        public IDisposable? BeginScope<TState>(TState state)
            where TState : notnull =>
            null;

        public bool IsEnabled(LogLevel logLevel) =>
            provider.IsEnabled(category, logLevel);

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel))
            {
                return;
            }

            provider.Write(category, logLevel, formatter(state, exception), state as IReadOnlyList<KeyValuePair<string, object?>>, exception);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;

/// <summary>
/// What the module has to say, to the Application event log under the source <c>OfficeHost.Shadow</c>, which
/// Install-OfficeHostShadow.ps1 creates, and to Trace: at most once a minute for each message, never throwing.
/// </summary>
static class TapLog
{
    const string source = "OfficeHost.Shadow";
    static Lock gate = new();
    static Dictionary<string, DateTime> written = [];
    static bool eventLog = true;

    public static void Write(string message)
    {
        try
        {
            var now = TapClock.UtcNow;
            var key = message.Length > 60 ? message.Substring(0, 60) : message;
            lock (gate)
            {
                if (written.TryGetValue(key, out var last) &&
                    now - last < TimeSpan.FromMinutes(1))
                {
                    return;
                }

                written[key] = now;
            }

            Trace.TraceWarning($"{source}: {message}");
            if (eventLog)
            {
                try
                {
                    EventLog.WriteEntry(source, message, EventLogEntryType.Warning);
                }
                catch (Exception)
                {
                    // no source, or no right to create one: Trace only
                    eventLog = false;
                }
            }
        }
        catch (Exception)
        {
            // never to SharePoint
        }
    }
}

using System;
using System.Threading;
using System.Web.Configuration;
using System.Web.Hosting;

/// <summary>
/// The module's state in an app domain: its settings, budget, queue and sender, made at the first request, with an
/// instance id of its own for the shadow. Registered with the hosting environment, so that when the app domain stops it
/// sends what waits and the final batch, within 5 seconds, without holding the recycle.
/// </summary>
sealed class TapRuntime : IRegisteredObject
{
    static Lock gate = new();
    static TapRuntime? current;
    static volatile bool read;
    long sequence;

    TapRuntime(TapSettings settings)
    {
        Settings = settings;
        Budget = new(settings.MaxInFlightBytes);
        Queue = new();
        ITapSink sink;
        if (settings.IsFolder)
        {
            sink = new FolderSink(settings.Sink!);
        }
        else
        {
            sink = new HttpSink(settings, Guid.NewGuid(), Environment.MachineName);
        }

        Sender = new(Queue, sink, settings.SamplePercent, TapLog.Write);
    }

    public TapSettings Settings { get; }

    public TapBudget Budget { get; }

    public TapQueue Queue { get; }

    public TapSender Sender { get; }

    /// <summary>
    /// Whether the module copies: enabled, and the sink answering.
    /// </summary>
    public bool Copying => !Sender.Open;

    /// <summary>
    /// The runtime of the app domain, or null when the module is disabled or not configured.
    /// </summary>
    public static TapRuntime? Get()
    {
        if (read)
        {
            return current;
        }

        lock (gate)
        {
            if (read)
            {
                return current;
            }

            try
            {
                var settings = TapSettings.Read(_ => WebConfigurationManager.AppSettings[_]);
                if (settings.Enabled)
                {
                    current = new(settings);
                    current.Sender.Start();
                    HostingEnvironment.RegisterObject(current);
                }
            }
            catch (Exception exception)
            {
                TapLog.Write($"The module is off: {exception.Message}");
                current = null;
            }

            read = true;
            return current;
        }
    }

    public long NextSequence() => Interlocked.Increment(ref sequence);

    public void Stop(bool immediate)
    {
        if (immediate)
        {
            HostingEnvironment.UnregisterObject(this);
            return;
        }

        // return at once: the flush runs on a thread of its own
        new Thread(() =>
        {
            try
            {
                Sender.Stop(TimeSpan.FromSeconds(5));
            }
            finally
            {
                HostingEnvironment.UnregisterObject(this);
            }
        })
        {
            IsBackground = true
        }.Start();
    }
}

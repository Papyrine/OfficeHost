/// <summary>
/// Sends the queue to the sink from a thread of its own, never from a request's: a batch whenever items wait, a heartbeat
/// (a batch of none, with the watermark) every few seconds otherwise. When the shadow is full it waits as the shadow
/// asks; when the sink fails three times in a row it is open: what waits becomes tombstones, the module copies nothing,
/// and it tries again at each heartbeat with backoff, until the sink answers.
/// </summary>
sealed class TapSender
{
    const int maximumItems = 50;
    const long maximumBytes = 32L * 1024 * 1024;
    TapQueue queue;
    ITapSink sink;
    Action<string> log;
    TimeSpan heartbeat;
    AutoResetEvent signal = new(false);
    ManualResetEvent stopped = new(false);
    Thread? thread;
    volatile bool stopping;
    volatile int samplePercent;
    volatile bool open;
    int failures;
    DateTime nextTry;

    // whether what waited became tombstones since the shadow last answered: once an outage
    bool unreachable;

    // the time Stop keeps to send what waits, after waiting for the requests in flight
    static TimeSpan sendTime = TimeSpan.FromSeconds(2);
    DateTime stopBy;

    public TapSender(TapQueue queue, ITapSink sink, int samplePercent, Action<string> log, TimeSpan? heartbeat = null)
    {
        this.queue = queue;
        this.sink = sink;
        this.samplePercent = samplePercent;
        this.log = log;
        this.heartbeat = heartbeat ?? TimeSpan.FromSeconds(5);
    }

    /// <summary>
    /// The percentage of documents to copy: the settings', or the shadow's when it sent one.
    /// </summary>
    public int SamplePercent => samplePercent;

    /// <summary>
    /// Whether the sink fails: the module copies nothing until it answers again.
    /// </summary>
    public bool Open => open;

    public void Start()
    {
        thread = new(Loop)
        {
            IsBackground = true,
            Name = "OfficeHost.Shadow sender"
        };
        thread.Start();
    }

    public void Signal() => signal.Set();

    /// <summary>
    /// Sends what waits, and the final batch, within the time given, then stops. The requests in flight, which the app
    /// domain drains as it stops, are waited for first, keeping time to send: the final batch is final only once none
    /// is, so that the shadow knows of the loss of the exchanges of those that end after it.
    /// </summary>
    public void Stop(TimeSpan flush)
    {
        stopBy = TapClock.UtcNow + flush - sendTime;
        stopping = true;
        signal.Set();
        stopped.WaitOne(flush);
    }

    void Loop()
    {
        try
        {
            while (!stopping)
            {
                signal.WaitOne(heartbeat);
                SendAll(false);
            }

            while (queue.InFlight > 0 &&
                   TapClock.UtcNow < stopBy)
            {
                signal.WaitOne(TimeSpan.FromMilliseconds(100));
            }

            SendAll(true);
        }
        catch (Exception exception)
        {
            log($"The sender stopped: {exception}");
        }
        finally
        {
            stopped.Set();
        }
    }

    /// <summary>
    /// Sends batches until none waits, or the sink fails: one batch, maybe empty, at least.
    /// </summary>
    public void SendAll(bool final)
    {
        var now = TapClock.UtcNow;
        if (open &&
            now < nextTry &&
            !final)
        {
            return;
        }

        do
        {
            var batch = queue.Take(maximumItems, maximumBytes, out var watermark, out var dropped);
            // final only with nothing to come: no request in flight either
            var last = final &&
                       queue.Count == 0 &&
                       queue.InFlight == 0;
            TapSent sent;
            try
            {
                sent = sink.Send(batch, watermark, dropped, last);
            }
            catch (Exception exception)
            {
                sent = new()
                {
                    Error = exception.Message
                };
            }

            if (sent.SamplePercent is { } sample)
            {
                samplePercent = Math.Max(0, Math.Min(100, sample));
            }

            if (sent.Sent)
            {
                foreach (var item in batch)
                {
                    item.Release();
                }

                if (open)
                {
                    log("The shadow answers again: copying again");
                }

                failures = 0;
                open = false;
                unreachable = false;
                continue;
            }

            queue.Return(batch, dropped);
            if (sent.RetryAfter is { } retryAfter)
            {
                // full or paused: the shadow answers, so wait as asked, the exchanges kept within the budget
                nextTry = TapClock.UtcNow + retryAfter;
                open = true;
                failures = 0;
                return;
            }

            failures++;
            nextTry = TapClock.UtcNow + TimeSpan.FromSeconds(Math.Min(60, Math.Pow(2, failures)));
            // once an outage, whether the shadow paused the module before it or not
            if (failures >= 3 &&
                !unreachable)
            {
                open = true;
                unreachable = true;
                queue.Bury("the shadow could not be reached");
                log($"The shadow cannot be reached ({sent.Error}): copying nothing until it answers");
            }

            return;
        }
        // the final pass sends all that waits, within the time of Stop
        while (queue.Count > 0 &&
               (final || !stopping));
    }
}

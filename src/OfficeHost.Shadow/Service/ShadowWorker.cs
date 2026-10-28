/// <summary>
/// Runs the <see cref="ShadowProcessor"/>: when a batch arrives, and every second for the exchanges waiting on the
/// watermarks of other servers.
/// </summary>
sealed class ShadowWorker(ShadowInbox inbox, ShadowProcessor processor, ILogger<ShadowWorker> log) :
    BackgroundService
{
    protected override async Task ExecuteAsync(Cancel cancel)
    {
        while (!cancel.IsCancellationRequested)
        {
            try
            {
                await inbox.Wait(TimeSpan.FromSeconds(1), cancel);
                await processor.Run(cancel);
            }
            catch (OperationCanceledException) when (cancel.IsCancellationRequested)
            {
                return;
            }
            catch (Exception exception)
            {
                var id = new EventId((int) ShadowOutcome.ShadowFailed, nameof(ShadowOutcome.ShadowFailed));
                log.LogWarning(id, exception, "The shadow failed a pass");
            }
        }
    }
}

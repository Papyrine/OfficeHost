namespace OfficeHost;

/// <summary>
/// Makes an OfficeHost application a shadow of SharePoint (docs/shadow.md): a deployment of its own, with its own
/// storage, that receives desktop Office's traffic with SharePoint from the OfficeHost.SharePoint module, replays it to
/// its OfficeHost, and logs how the answers compare.
/// </summary>
public static class OfficeHostShadowExtensions
{
    /// <summary>
    /// Registers the shadow, with the options of the <c>OfficeHostShadow</c> section then <paramref name="configure"/>.
    /// <see cref="OfficeHostExtensions.AddOfficeHost"/> and an <see cref="IShadowDocuments"/> (a singleton) must be
    /// registered too; an <see cref="IShadowContent"/> may be, to start the documents Office opens from its cache.
    /// Replaces the <see cref="TimeProvider"/>: OfficeHost replays each exchange at the time SharePoint received it.
    /// </summary>
    public static IServiceCollection AddOfficeHostShadow(this IServiceCollection services, Action<OfficeHostShadowOptions>? configure = null)
    {
        var options = services.AddOptions<OfficeHostShadowOptions>().BindConfiguration("OfficeHostShadow");
        if (configure != null)
        {
            options.Configure(configure);
        }

        services.Replace(ServiceDescriptor.Singleton<TimeProvider>(new ShadowClock()));
        services.TryAddSingleton<ShadowInbox>();
        services.TryAddSingleton<ServiceFiles>();
        services.TryAddSingleton<InProcessTarget>();
        services.TryAddSingleton<ShadowProcessor>();
        services.AddHostedService<ShadowWorker>();
        return services;
    }

    /// <summary>
    /// Maps the endpoint the module in SharePoint posts to: <c>{pattern}/exchanges</c>.
    /// </summary>
    public static RouteGroupBuilder MapOfficeHostShadow(this IEndpointRouteBuilder endpoints, string pattern = "/officehost-shadow")
    {
        var group = endpoints.MapGroup(pattern);
        group.MapPost("/exchanges", ShadowIngest.Handle);
        return group;
    }
}

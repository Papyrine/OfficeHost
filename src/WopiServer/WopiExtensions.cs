namespace WopiServer;

/// <summary>
/// Adds the WOPI endpoints to an ASP.NET Core application.
/// </summary>
public static class WopiExtensions
{
    /// <summary>
    /// Registers the WOPI services. An <see cref="IWopiStorage"/> must be registered too; <see cref="IWopiLockStore"/>
    /// defaults to <see cref="MemoryLockStore"/>, <see cref="IWopiTokens"/> to tokens signed with
    /// <see cref="WopiOptions.SigningKey"/>, and PutUserInfo is supported when an <see cref="IWopiUserInfoStore"/> is.
    /// These can have any lifetime, as a scoped storage over a database context.
    /// </summary>
    public static IServiceCollection AddWopi(this IServiceCollection services, Action<WopiOptions> configure)
    {
        services.Configure(configure);
        services.AddHttpClient(WopiOptions.HttpClientName);
        services.TryAddSingleton(TimeProvider.System);
        services.TryAddSingleton<IWopiLockStore, MemoryLockStore>();
        services.TryAddSingleton<IWopiTokens, SignedWopiTokens>();
        services.TryAddSingleton<WopiDiscovery>();
        services.TryAddSingleton<ProofKeys>();
        // transient, so that they take the lifetime of the storage and the stores rather than capture them
        services.TryAddTransient<WopiHandler>();
        services.TryAddTransient(
            _ => new WopiLauncher(
                _.GetRequiredService<IWopiStorage>(),
                _.GetRequiredService<IWopiTokens>(),
                _.GetRequiredService<IOptions<WopiOptions>>(),
                _.GetRequiredService<WopiDiscovery>()));
        return services;
    }

    /// <summary>
    /// Maps the WOPI endpoints under <c>/wopi/files</c>, as WOPI clients expect. The WOPISrc of a file is
    /// <see cref="WopiOptions.BaseUrl"/> followed by <c>/wopi/files/{fileId}</c>.
    /// </summary>
    public static RouteGroupBuilder MapWopi(this IEndpointRouteBuilder endpoints)
    {
        var services = endpoints.ServiceProvider;
        if (!services.GetRequiredService<IServiceProviderIsService>().IsService(typeof(IWopiStorage)))
        {
            throw new InvalidOperationException($"No {nameof(IWopiStorage)} is registered: register one alongside {nameof(AddWopi)}");
        }

        // fails now rather than on the first request when the signing key is missing
        using (var scope = services.CreateScope())
        {
            scope.ServiceProvider.GetRequiredService<IWopiTokens>();
        }

        var options = services.GetRequiredService<IOptions<WopiOptions>>().Value;
        var serverVersion = options.ServerVersion ?? typeof(WopiExtensions).Assembly.GetName().Version?.ToString() ?? "1.0";
        var machineName = Environment.MachineName;
        var log = services.GetRequiredService<ILoggerFactory>().CreateLogger("Wopi");
        var group = endpoints.MapGroup("/wopi");
        group.AddEndpointFilter(
            async (context, next) =>
            {
                var http = context.HttpContext;
                // [MS-WOPI] 2.2.1 requires these on all WOPI responses
                http.Response.Headers["X-WOPI-ServerVersion"] = serverVersion;
                http.Response.Headers["X-WOPI-MachineName"] = machineName;
                try
                {
                    return await next(context);
                }
                catch (BadHttpRequestException exception) when (exception.StatusCode == StatusCodes.Status413PayloadTooLarge)
                {
                    return WopiResult.Text("The file is too large", StatusCodes.Status413PayloadTooLarge);
                }
                catch (Exception exception) when (!http.RequestAborted.IsCancellationRequested)
                {
                    log.LogError(exception, "Unexpected failure of {Method} {Path}", http.Request.Method, http.Request.Path);
                    return WopiResult.ServerError("Internal error");
                }
            });

        group.MapGet("/files/{fileId}", (HttpContext context, string fileId, WopiHandler handler) => handler.CheckFileInfo(context, fileId));
        group.MapGet("/files/{fileId}/contents", (HttpContext context, string fileId, WopiHandler handler) => handler.GetFile(context, fileId));
        group.MapPost("/files/{fileId}", (HttpContext context, string fileId, WopiHandler handler) => handler.FilesPost(context, fileId));
        group.MapPost("/files/{fileId}/contents", (HttpContext context, string fileId, WopiHandler handler) => handler.PutFile(context, fileId));
        // the folder operations of [MS-WOPI] 3.3.5.4 and up, and the ecosystem ones, are not supported
        group.Map("/{**path}", () => WopiResult.NotImplemented("Unsupported WOPI endpoint"));
        return group;
    }
}

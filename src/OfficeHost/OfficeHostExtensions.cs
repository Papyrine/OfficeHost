namespace OfficeHost;

/// <summary>
/// Adds the WOPI endpoints to an ASP.NET Core application.
/// </summary>
public static class OfficeHostExtensions
{
    /// <summary>
    /// Registers the WOPI services. An <see cref="IOfficeStorage"/> must be registered too; <see cref="IOfficeLockStore"/>
    /// defaults to <see cref="MemoryLockStore"/>, <see cref="IFssHttpStore"/> to <see cref="MemoryFssHttpStore"/>,
    /// <see cref="IOfficeTokens"/> to tokens signed with
    /// <see cref="OfficeHostOptions.SigningKey"/>, and PutUserInfo is supported when an <see cref="IOfficeUserInfoStore"/> is.
    /// These can have any lifetime, as a scoped storage over a database context.
    /// </summary>
    public static IServiceCollection AddOfficeHost(this IServiceCollection services, Action<OfficeHostOptions> configure)
    {
        services.Configure(configure);
        services.AddHttpClient(OfficeHostOptions.HttpClientName);
        services.TryAddSingleton(TimeProvider.System);
        services.TryAddSingleton<IOfficeLockStore, MemoryLockStore>();
        services.TryAddSingleton<IFssHttpStore, MemoryFssHttpStore>();
        services.TryAddSingleton<IOfficeTokens, SignedOfficeTokens>();
        services.TryAddSingleton<WopiDiscovery>();
        services.TryAddSingleton<ProofKeys>();
        // transient, so that they take the lifetime of the storage and the stores rather than capture them
        services.TryAddTransient<WopiHandler>();
        services.TryAddTransient<FssHttpCells>();
        services.TryAddTransient<CoauthSessions>();
        services.TryAddTransient<CellStorageService>();
        services.TryAddTransient<FssHttpHandler>();
        services.TryAddTransient(
            _ => new OfficeLauncher(
                _.GetRequiredService<IOfficeStorage>(),
                _.GetRequiredService<IOfficeTokens>(),
                _.GetRequiredService<IOptions<OfficeHostOptions>>(),
                _.GetRequiredService<WopiDiscovery>()));
        return services;
    }

    /// <summary>
    /// Maps the WOPI endpoints under <c>/wopi/files</c>, as WOPI clients expect. The WOPISrc of a file is
    /// <see cref="OfficeHostOptions.BaseUrl"/> followed by <c>/wopi/files/{fileId}</c>.
    /// </summary>
    public static RouteGroupBuilder MapOfficeHost(this IEndpointRouteBuilder endpoints)
    {
        var services = endpoints.ServiceProvider;
        if (!services.GetRequiredService<IServiceProviderIsService>().IsService(typeof(IOfficeStorage)))
        {
            throw new InvalidOperationException($"No {nameof(IOfficeStorage)} is registered: register one alongside {nameof(AddOfficeHost)}");
        }

        // fails now rather than on the first request when the signing key is missing
        using (var scope = services.CreateScope())
        {
            scope.ServiceProvider.GetRequiredService<IOfficeTokens>();
        }

        var options = services.GetRequiredService<IOptions<OfficeHostOptions>>().Value;
        var serverVersion = options.ServerVersion ?? typeof(OfficeHostExtensions).Assembly.GetName().Version?.ToString() ?? "1.0";
        var machineName = Environment.MachineName;
        var log = services.GetRequiredService<ILoggerFactory>().CreateLogger("OfficeHost");
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

    /// <summary>
    /// Maps the sites desktop Office edits and coauthors files through, with [MS-FSSHTTP] as SharePoint has it, under
    /// <see cref="OfficeHostOptions.DesktopPath"/>: <c>{DesktopPath}/{token}/{name}</c> is the URL of a file, and
    /// <see cref="OfficeLauncher.LaunchDesktop"/> the link that opens it. The coauthoring state is kept in the
    /// <see cref="IFssHttpStore"/>, and the lock of desktop editing in the <see cref="IOfficeLockStore"/>, so that WOPI
    /// clients and desktop Office do not edit a file at the same time.
    /// </summary>
    public static RouteGroupBuilder MapFssHttp(this IEndpointRouteBuilder endpoints)
    {
        var services = endpoints.ServiceProvider;
        if (!services.GetRequiredService<IServiceProviderIsService>().IsService(typeof(IOfficeStorage)))
        {
            throw new InvalidOperationException($"Register an {nameof(IOfficeStorage)} before mapping the desktop sites");
        }

        var path = services.GetRequiredService<IOptions<OfficeHostOptions>>().Value.DesktopPath;
        var group = endpoints.MapGroup(path);
        string[] verbs = ["OPTIONS", "POST", "GET", "HEAD"];
        group.MapMethods("/{token}", verbs, (HttpContext context, string token, FssHttpHandler handler) => handler.Handle(context, token, null));
        group.MapMethods("/{token}/{**rest}", verbs, (HttpContext context, string token, string? rest, FssHttpHandler handler) => handler.Handle(context, token, rest));
        return group;
    }
}

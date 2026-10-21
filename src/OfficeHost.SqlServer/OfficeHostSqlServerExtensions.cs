namespace OfficeHost;

public static class OfficeHostSqlServerExtensions
{
    /// <summary>
    /// Registers the SQL Server stores: <see cref="SqlServerStorage"/> as the <see cref="IOfficeStorage"/>, with
    /// <see cref="SqlServerLockStore"/>, <see cref="SqlServerFssHttpStore"/> and <see cref="SqlServerUserInfoStore"/>,
    /// in place of the defaults of <see cref="OfficeHostExtensions.AddOfficeHost"/>, before or after it. The tables
    /// are created by OfficeHost.SqlServer.Install.
    /// </summary>
    public static IServiceCollection AddOfficeHostSqlServer(this IServiceCollection services, string connectionString, string schema = "dbo")
    {
        services.AddSingleton(new SqlServerStorage(connectionString, schema));
        services.AddSingleton<IOfficeStorage>(_ => _.GetRequiredService<SqlServerStorage>());
        services.AddSingleton<IOfficeLockStore>(_ => new SqlServerLockStore(connectionString, _.GetService<TimeProvider>() ?? TimeProvider.System, schema));
        services.AddSingleton<IFssHttpStore>(new SqlServerFssHttpStore(connectionString, schema));
        services.AddSingleton<IOfficeUserInfoStore>(new SqlServerUserInfoStore(connectionString, schema));
        return services;
    }
}

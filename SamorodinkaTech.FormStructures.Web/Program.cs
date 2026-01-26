using SamorodinkaTech.FormStructures.Web.Services;
using Serilog;

var builder = WebApplication.CreateBuilder(args);

var logDir = Path.Combine(builder.Environment.ContentRootPath, "logs");
Directory.CreateDirectory(logDir);

builder.Host.UseSerilog((context, services, loggerConfiguration) =>
{
    loggerConfiguration
        .ReadFrom.Configuration(context.Configuration)
        .ReadFrom.Services(services)
        .Enrich.FromLogContext()
        .WriteTo.Console()
        .WriteTo.File(
            path: Path.Combine(logDir, "app-.log"),
            rollingInterval: RollingInterval.Day,
            retainedFileCountLimit: 14,
            shared: true,
            flushToDiskInterval: TimeSpan.FromSeconds(1));
});

// Add services to the container.
builder.Services.AddRazorPages();

builder.Services.AddOptions<StorageOptions>()
    .Bind(builder.Configuration.GetSection("Storage"));

builder.Services.AddSingleton<ExcelFormParser>();
builder.Services.AddSingleton<FormStorage>();
builder.Services.AddSingleton<FormDataStorage>();

var app = builder.Build();

// Housekeeping: remove stale pending schema uploads so storage does not grow unbounded.
// Can also be run as a one-off command: `dotnet run --project ... -- --cleanup-pending`.
{
    using var scope = app.Services.CreateScope();
    var storage = scope.ServiceProvider.GetRequiredService<FormStorage>();

    if (args.Contains("--cleanup-pending", StringComparer.OrdinalIgnoreCase))
    {
        var deleted = storage.CleanupOldPendingUploads(TimeSpan.FromDays(7));
        app.Logger.LogInformation("Cleanup complete. Deleted {Count} pending uploads.", deleted);
        return;
    }

    try
    {
        var deleted = storage.CleanupOldPendingUploads(TimeSpan.FromDays(7));
        if (deleted > 0)
        {
            app.Logger.LogInformation("Startup cleanup deleted {Count} pending uploads.", deleted);
        }
    }
    catch (Exception ex)
    {
        app.Logger.LogWarning(ex, "Startup cleanup for pending uploads failed.");
    }
}

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
    app.UseHttpsRedirection();
}

app.UseRouting();

app.UseAuthorization();

app.MapStaticAssets();
app.MapRazorPages()
   .WithStaticAssets();

app.Run();

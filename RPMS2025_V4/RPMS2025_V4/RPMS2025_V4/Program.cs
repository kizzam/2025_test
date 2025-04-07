using RPMS2025_V4.Client.Pages;
using RPMS2025_V4.Components;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using RPMS2025_V4.Data;
using RPMS2025_V4.Repositories;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddDbContextFactory<RPMSdBContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("RPMSConnection") ?? throw new InvalidOperationException("Connection string 'RPMSConnection' not found.")));

builder.Services.AddScoped<IBloodlineRepository, BloodlineRepository>();

builder.Services.AddQuickGridEntityFrameworkAdapter();

builder.Services.AddDatabaseDeveloperPageExceptionFilter();

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddInteractiveWebAssemblyComponents();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseWebAssemblyDebugging();
}
else
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
    app.UseMigrationsEndPoint();
}

app.UseHttpsRedirection();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(typeof(RPMS2025_V4.Client._Imports).Assembly);

app.Run();

using RPMS2025_WEB_2512.Client.Pages;
using RPMS2025_WEB_2512.Components;
using Microsoft.EntityFrameworkCore;
using RPMS2025_WEB_2512.Data;
using RPMS2025_WEB_2512.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Syncfusion.Blazor;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
//Syncfusion
builder.Services.AddRazorComponents();
builder.Services.AddSyncfusionBlazor();

//
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddInteractiveWebAssemblyComponents();

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();

// Get configuration
var configuration = builder.Configuration;

// Add the DbContext with the connection string
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("MyConnection")));

builder.Services.AddDbContext<BloodLineDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("MyConnection")));

// Register the connection string
IServiceCollection serviceCollection = builder.Services.AddSingleton(builder.Configuration.GetConnectionString("MyConnection"));

//KM
builder.Services.AddScoped<BloodlineServices>();
builder.Services.AddScoped<BloodlineServices_New>();

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
}

app.UseHttpsRedirection();

app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode()
    .AddInteractiveWebAssemblyRenderMode()
    .AddAdditionalAssemblies(typeof(RPMS2025_WEB_2512.Client._Imports).Assembly);

app.Run();

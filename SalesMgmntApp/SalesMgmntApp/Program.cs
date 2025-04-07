using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.EntityFrameworkCore;
using SalesMgmntApp.Data;
using Syncfusion.Blazor;


var builder = WebApplication.CreateBuilder(args);

var ConnectionString = builder.Configuration.GetConnectionString("BirdMgmntDbConnection")
    ?? throw new InvalidOperationException("Connection 'BirdMgmntConnection' NOT Found");

builder.Services.AddDbContext<BirdsMgmntDbContext>(
    options => options.UseSqlServer(ConnectionString));

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddSingleton<WeatherForecastService>();

builder.Services.AddSyncfusionBlazor();

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
    "OTQ0NzQyQDMyMzAyZTM0MmUzMG41ZGdReWRqQ1UvalJuSEtvakgxZmtrV1NOKytXVC9MUmxEMEswbm1RY0k9");

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();

app.UseRouting();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.Run();

//1.06.26 runs ok
//write code in rpms to export data in format for input seed data here
//


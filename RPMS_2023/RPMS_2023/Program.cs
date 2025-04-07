using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.Extensions.Configuration;
using RPMS_2023.Data;
using Syncfusion.Blazor;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddSingleton<WeatherForecastService>();

builder.Services.AddSyncfusionBlazor();

var configuration = new ConfigurationBuilder()
                             .AddJsonFile("appsettings.json", false, true)
                             .AddEnvironmentVariables()
                             .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT")}.json", true, true)
                             .Build();

var SqlConnectionConfiguration = new SqlConnectionConfiguration(configuration.GetConnectionString("SqlDBContext"));

//builder.Services.AddDbContext<LibraryContext>(option =>
   //             option.UseSqlSe       GetConnectionString("LibraryDatabase")));

builder.Services.AddSingleton(SqlConnectionConfiguration);

builder.Services.AddScoped<IColourService, ColourService>();

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("NzQwMzczQDMyMzAyZTMzMmUzMG1CajBieUlySmdkeGhSdEhqT2JuOHpONGNkeThkOVU1N1hlbzVsQTZhcU09");

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

//Wheres Configuration services???
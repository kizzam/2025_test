using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using RPMS.Plugins.InMemory;
using RPMS.UseCases.Pigeons.Interfaces;
using RPMS.UseCases.Pigeons;
using RPMS.UseCases.PluginInterfaces;
using RPMS.WebApp.Data;
//using RPMS.UseCases.Pigeons;


var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddSingleton<WeatherForecastService>();

builder.Services.AddSingleton<IPigeonRepository, PigeonRepository>();

builder.Services.AddTransient<IViewPigeonsByRingNoUseCase, ViewPigeonsByRingNoUseCase>();
builder.Services.AddTransient<IAddPigeonUseCase, AddPigeonUseCase>();

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

// End os Session
// 1:13:17
// 1:53:15
// 2:09:43
// 2:18:14 works
// 02:21:15 works
// 02:26:20 works
// 02:27:19
//nb: not working due to need to find code for iAddPigeonUseCase < no executeasync function
//
//


// Progress
// 2:00:40 Button now has edit
// 2:03:46 - 2:03:58 is now new component with EditInventory???

// Starts at 2:04:04
// 2:04:45
// Problem EditPigeon use case not defined
// OK 2:09:43
// 02:11:44
// 02:13:44
// 02:21:15


//

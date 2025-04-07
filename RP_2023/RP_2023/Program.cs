global using System.Data.SqlClient;
global using Dapper;
global using Dapper.Contrib.Extensions;
global using System.Linq.Expressions;
global using System.Data;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using RP_2023.Data;
using Microsoft.EntityFrameworkCore;
using System.Configuration;
using RP_2023.Models;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddSingleton<WeatherForecastService>();

builder.Services.AddSingleton<DapperRepository<Colour2>>(s =>
    new DapperRepository<Colour2>(
        builder.Configuration.GetConnectionString("RPConnectionString")));

builder.Services.AddSingleton<DapperRepository<Colour>>(s =>
    new DapperRepository<Colour>(
        builder.Configuration.GetConnectionString("RPConnectionString")));

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

// Read the connection string from the appsettings.json file
// Set the database connection for the EndtoEndContext
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");


app.Run();

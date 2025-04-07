using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using MyBlazorWAapp_SF;
using Syncfusion.Blazor;


var builder = WebAssemblyHostBuilder.CreateDefault(args);

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("NjMzOTI5QDMyMzAyZTMxMmUzMEdiMyttSlpOc2NPdmtJVThTeWl3MVJWZEYvRVo0cDFxL2YrWGdmMXNUVUE9");

builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });

builder.Services.AddSyncfusionBlazor();


await builder.Build().RunAsync();

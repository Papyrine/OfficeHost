using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Morph;

// the editor: MorphWopi opens the file that Morph.Wopi writes into the page, through the WOPI endpoints of the sample
var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<MorphWopi>("#app");
builder.Services.AddScoped(
    _ => new HttpClient
    {
        BaseAddress = new(builder.HostEnvironment.BaseAddress)
    });
builder.Services.AddMorph();
await builder.Build().RunAsync();

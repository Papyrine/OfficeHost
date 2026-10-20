The host side of the Office document protocols for ASP.NET Core, formerly the WopiServer package. Add it to a web
application, and WOPI clients such as Collabora Online, ONLYOFFICE and Microsoft 365 for the web open, edit and save the
files of any storage the application has.

```csharp
using OfficeHost;

builder.Services.AddSingleton<IOfficeStorage>(new FileSystemStorage("/var/documents"));
builder.Services.AddOfficeHost(options => options.SigningKey = builder.Configuration["OfficeHost:SigningKey"]);

var app = builder.Build();
app.MapOfficeHost();
```

`OfficeLauncher` gives the URL of the editor and the access token for a file. See https://github.com/Papyrine/OfficeHost
for the documentation and a sample application.

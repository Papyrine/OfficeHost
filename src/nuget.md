The host side of the Office document protocols for ASP.NET Core, formerly the WopiServer package. Add it to a web
application, and the files of any storage the application has open, edit and save in WOPI clients such as Collabora
Online, ONLYOFFICE and Microsoft 365 for the web, and, in preview, in desktop Word, Excel and PowerPoint through the cell
storage service of SharePoint (MS-FSSHTTP), with coauthoring.

```csharp
using OfficeHost;

builder.Services.AddSingleton<IOfficeStorage>(new FileSystemStorage("/var/documents"));
builder.Services.AddOfficeHost(options => options.SigningKey = builder.Configuration["OfficeHost:SigningKey"]);

var app = builder.Build();
app.MapOfficeHost();
// desktop Office
app.MapFssHttp();
```

`OfficeLauncher` gives the URL of the editor and the access token for a file, or the `ms-word:` (`ms-excel:`,
`ms-powerpoint:`) link that opens it in desktop Office. See https://github.com/Papyrine/OfficeHost
for the documentation and a sample application.

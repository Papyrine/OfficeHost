The host side of the Web Application Open Platform Interface (WOPI) for ASP.NET Core. Add it to a web application, and
WOPI clients such as Collabora Online, ONLYOFFICE and Microsoft 365 for the web open, edit and save the files of any
storage the application has.

```csharp
builder.Services.AddSingleton<IWopiStorage>(new FileSystemStorage("/var/documents"));
builder.Services.AddWopi(options => options.SigningKey = builder.Configuration["Wopi:SigningKey"]);

var app = builder.Build();
app.MapWopi();
```

`WopiLauncher` gives the URL of the editor and the access token for a file. See https://github.com/Papyrine/WopiServer
for the documentation and a sample application.

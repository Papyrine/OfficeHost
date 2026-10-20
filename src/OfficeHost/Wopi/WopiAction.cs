/// <summary>
/// An action of a WOPI client, as its discovery lists it.
/// </summary>
/// <param name="App">The app that offers the action, as Word or writer.</param>
/// <param name="Name">The action, as edit or view.</param>
/// <param name="Extension">The file extension it applies to, without the dot.</param>
/// <param name="UrlSrc">The URL template of the action.</param>
record WopiAction(string App, string Name, string Extension, string UrlSrc);
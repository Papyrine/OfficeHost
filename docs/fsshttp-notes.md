# FSSHTTP: where Word differs from the specs

Each entry gives the spec section, what Word did, and the evidence: a run of the spike (docs/fsshttp-spike.md), whose
logs are in `src/WopiServer.Sample/spike-logs`. The client is Word 16.0.20326 (Microsoft 365) on Windows 11 10.0.26200.


## Discovery

### No FrontPage RPC despite `MS-Author-Via: MS-FP/4.0`

- **Spec:** [MS-WDVME] 3.2.5.1 and 4.1.2 have a FrontPage-compliant server advertise `MS-Author-Via: MS-FP/4.0,DAV`.
  [MS-FPSE] defines `_vti_inf.html` and `author.dll` for clients that use it.
- **Word:** in runs 1, 2 and 5 every OPTIONS answer had `MS-Author-Via: MS-FP/4.0,DAV`, `DocumentManagementServer`
  and `MicrosoftSharePointTeamServices`. Word sent no FrontPage request at all (`_vti_inf.html`, `shtml.dll`,
  `author.dll`), and no `webs.asmx` either. It used WebDAV only: OPTIONS, HEAD, LOCK, GET, PROPFIND, UNLOCK.

### Undocumented OPTIONS request headers

- **Spec:** nothing found in [MS-WDVME], [MS-WEBDAVE], [MS-OFBA] or [MS-FSSHTTP].
- **Word:** every OPTIONS of runs 1 to 5 sent these headers:

  ```
  User-Agent: Microsoft Office Word 2014 (16.0.20326) Windows NT 10.0
  Authorization: Bearer
  X-Office-Major-Version: 16
  X-MS-CookieUri-Requested: t
  X-FeatureVersion: 1
  X-MS-AuthDomainSupport: True
  Accept-Auth: badger,Wlid1.1,Bearer,Basic,NTLM,Digest,Kerberos,Negotiate,Nego2
  X-MSGETWEBURL: t
  X-IDCRL_ACCEPTED: t
  ```

  Some OPTIONS used the user agent `Microsoft Office Word` alone. In each run, one OPTIONS replaced
  `X-MSGETWEBURL` with `X-IDCRL_OPTIONS: force-auth-challenge` and `IgnoreCookieAuthentication: t`.
- **Open:** the server's answer to `X-MSGETWEBURL` and `X-MS-CookieUri-Requested`. It needs a SharePoint Online
  capture.

### FSSHTTP is negotiated on OPTIONS: X-MSFSSHTTP and X-MSWEBURL

- **Spec:** [MS-FSSHTTP] 1.7 (Versioning and Capability Negotiation) is "None", and 1.5 only says where the endpoint
  is. Neither header is in [MS-FSSHTTP], [MS-WDVME], [MS-WEBDAVE] or [MS-OFBA].
- **SharePoint Server 16.0.0.5522, signed in:**
  - every OPTIONS answer had `X-MSFSSHTTP: 1.3`;
  - the OPTIONS that carried `X-MSGETWEBURL: t` also got `X-MSWEBURL: /`, the server-relative path of the site;
  - Word then POSTed to `<web url>/_vti_bin/cellstorage.svc/CellStorageService`, with no FrontPage or `webs.asmx`
    request.
- **The spike:**
  - without these headers (runs 1 to 5), Word used plain WebDAV;
  - with them (run 6, `X-MSWEBURL: /n`), Word POSTed to `/n/_vti_bin/cellstorage.svc/CellStorageService` the same
    request it sent SharePoint.
- **Before sign-in** SharePoint sends neither (see below), so a host must not answer OPTIONS with them unless the
  request is authenticated, or Word would try FSSHTTP without credentials.

### What Word does when the Cell sub-requests fail

- **Spec:** [MS-FSSHTTP] 2.2.5.6 has `RequestNotSupported` for a sub-request the server does not support.
- **Word,** with `RequestNotSupported` on the Cell sub-requests and success on the rest, sent `ExitCoauthoring` twice.
  Then:
  - **Run 6 (Windows sign-in, `/n`):** it showed "Sorry, we couldn't open '<url>'", with no GET, HEAD or LOCK.
  - **Run 7 (token in the path, anonymous, `/t/{token}`):** it fell back. It did a GET of the document, then the old
    discovery at the root of the host (`_vti_inf.html`, the FrontPage RPC, HEAD, `webs.asmx` `WebUrlFromPageUrl`,
    `workflow.asmx` `GetWorkflowDataForItem`). Then it opened the document read-only, in Protected View.
- **Why they differ:** unknown. The difference between the runs is the sign-in, and the site path.

### The first request differs from the example of MS-FSSHTTP 4.1.1

- **Spec:** [MS-FSSHTTP] 4.1.1 has Coauth, SchemaLock, three Cell, ServerTime and WhoAmI.
- **Word 16.0.20326,** to SharePoint and to the spike alike:
  - **Extra sub-requests:** three EditorsTable `UpdateEditorMetadata` (`OnSuccess` the join; keys `SupportsRename`,
    `Capabilities`, `Platform`, base64 values), `GetDocMetaInfo` (`OnExecute` WhoAmI), `GetVersions`, and a fourth,
    27-byte Cell sub-request with no dependency.
  - **`Timeout`:** `2580`.
  - **`Request` attributes:** `UseResourceID="true"` (plus `ResourceID` when Word knows it from an earlier open),
    `UserAgent` (a GUID), `UserAgentClient="msword"`, `UserAgentPlatform="win"`, `Build`, `MetaData="1031"`.
  - **HTTP headers:** besides SOAPAction, `Application`, `FileExtension`, `OSName`, `Scenario: OpenFile_Open`,
    `ScenarioType`, `X-ClientRing` and `X-RequestStats`.
  - **Encoding:** the body is MTOM, but the Cell data is inline base64 (with `BinaryDataSize`), not an XOP part.

### SharePoint's response to it

- **Spec:** [MS-FSSHTTP] 4.1.2 has `ResponseVersion MinorVersion="0"`, `WebUrlIsEncoded="false"`, and
  `UrlIsEncoded="true"`.
- **SharePoint Server 16.0.0.5522:**
  - **Envelope:** `MinorVersion="3"`; `ResponseCollection WebUrl="https://<host>"` without `WebUrlIsEncoded`.
  - **`Response`:** `Url` unencoded (with spaces) and without `UrlIsEncoded`; it echoes `ResourceID`.
  - **The join:** `TransitionID` equals the `ResourceID`.
  - **Cell sub-responses:** `ServerCorrelationId` (the `SPRequestGuid`). Their data is inline base64 for some and an
    XOP part (`Content-Transfer-Encoding: binary`) for others, in the same response.
  - **WhoAmI:** `UserLogin` is a claims login (`i:0#.w|<domain>\<user>`); `UserSIPAddress` is empty.

### sharedaccess.asmx is polled from the open

- **Spec:** [MS-SHDACCWS] `IsOnlyClient`.
- **Word:**
  - **Timing:** right after the cellstorage open, it POSTs `/_vti_bin/sharedaccess.asmx`
    (`SOAPAction: "http://schemas.microsoft.com/sharepoint/soap/IsOnlyClient"`), every 8 seconds for a minute and
    then every 30 seconds.
  - **Body:** `<IsOnlyClient><id>` is the `ResourceID`, and SharePoint answered `true`.
  - **HTTP headers:** `Scenario: PollingSubscriptionNotConnected_AmIAlone` and
    `X-RequestStats: Operation=AmIAlone`.


### Required columns block the save before any upload

- **SharePoint Server, a library with required columns:** after an edit, Word's save showed "REQUIRED PROPERTIES: To
  save to the server, correct the invalid or missing required properties", and sent no cellstorage request. Word
  checks the library's columns, from what it learnt at open (GetDocMetaInfo or the metadata partition), before it
  uploads.


## Authentication

### MS-OFBA is blocked by default, for SharePoint too

- **Spec:** [MS-OFBA] 2.2.1 has the client identify itself with `X-FORMS_BASED_AUTH_ACCEPTED: t` (or a listed user
  agent). On a 403 with `X-FORMS_BASED_AUTH_REQUIRED` (2.2.2), the client MUST navigate to the login page.
- **Word:** it sent no `X-FORMS_BASED_AUTH_ACCEPTED`. On the 403 of run 3 it did not open the login page. It showed
  "Office has blocked this content because it uses a sign-in method that may be insecure. You can choose to not block
  such content in the future by changing your settings in the Office Trust Center."
- **SharePoint Server 16.0.0.5522 (forms-based sign-in), captured with `src/fsshttp-capture.ps1`:** Word sent the same
  OPTIONS. SharePoint answered 403 with `X-Forms_Based_Auth_Required`, `X-Forms_Based_Auth_Return_Url`,
  `X-MSDAVEXT_Error: 917656; Access denied. Before opening files in this location, you must first browse to the web
  site and select the option to login automatically.` and `X-MS-InvokeApp: 1; RequireReadOnly`. Word showed the same
  dialog and sent nothing more.
- **The setting:** File > Options > Trust Center > Trust Center Settings > Form-based Sign-in. Its Group Policy is
  "Control how Office handles form-based sign-in prompts". By default it blocks every prompt; it can be set to allow
  listed hosts.

### Word's forms sign-in dialog is the IE engine

- **Spec:** [MS-OFBA] 3.2 leaves the login dialog to the client.
- **Word:** with forms-based sign-in allowed for the host ("Ask me what to do for each host", then Yes), Word opened
  the login page in a `TridentContainer` window. That is the IE engine (MSHTML), and its requests go through WinINet
  inside the Word process. They carry the IE `Accept` header (`image/gif, image/jpeg, …, application/x-ms-xbap`).
- **SharePoint Server's login page** lists "Windows Authentication" and "ADFS". Both failed on a PC that is not on the
  SharePoint domain:
  - **Windows:** `POST /_login/default.aspx` then 302 to `/_windows/default.aspx`. The 401 `WWW-Authenticate: NTLM`
    led to an NTLM type 2 challenge, and the dialog stopped there.
  - **ADFS:** 302 to `/_trust/default.aspx?trust=ADFS`, 302 to `https://<adfs>/adfs/ls/?wa=wsignin1.0&wtrealm=urn:…`.
    ADFS answered with a 302 to `/adfs/ls/wia`, its Windows-integrated endpoint, because it treats IE as a browser
    that can sign in with Windows. Chrome gets the ADFS password form instead.

  In both cases the dialog showed "Can't reach this page" (titled `javascript:` the first time) and never asked for a
  password.
- **What it means for a host:** forms-based sign-in only works in Word if the login page works in the IE engine. Its
  identity provider must not send IE to Windows sign-in on PCs outside the domain.

### Word calls the SharePoint REST v2 API before sign-in

- **Spec:** none: `/_api/v2.0` is the SharePoint REST API (OneDrive API shape), not an Open Specification.
- **Word:** after the 403 of the OPTIONS, it sent `GET /_api/v2.0/shares/u!<base64url of the document URL>/driveitem/?$select=viewpoint`
  through WebIO (WinHTTP). SharePoint Server answered `401 WWW-Authenticate: NTLM`.

### SharePoint sends no capability headers before sign-in

- **SharePoint Server 16.0.0.5522:** the 403 answer to an unauthenticated OPTIONS has no `DAV`, `MS-Author-Via`,
  `Allow` or web URL; only `MicrosoftSharePointTeamServices` and the sign-in headers. So [MS-WDVME] 3.2.5.1 (which
  advertises capabilities on OPTIONS) applies only after sign-in, and so does any answer to `X-MSGETWEBURL`.

### Basic over HTTPS is blocked by default

- **Spec:** none: Basic is [RFC7617], and the `Accept-Auth` Word sends lists `Basic`.
- **Word:** on the 401 `WWW-Authenticate: Basic` of run 4 it showed no credential prompt. It showed "Microsoft Office
  has blocked access to "https://localhost:5001/b/" because the source uses a sign-in method that may be unsecure."

### Windows auth is silent

- **Word:** on the 401 `WWW-Authenticate: Negotiate` of run 5, it authenticated with no prompt and went on as an
  anonymous client would.


## WebDAV

### LOCK and UNLOCK around a read-only open

- **Word:** in Protected View (all runs on `https://localhost`, which is in the Internet zone), it took an exclusive
  LOCK (`Timeout: Second-600`, owner `AzureAD\<user>`), did a GET and a PROPFIND, and then an UNLOCK within a second.
  After "Enable Editing" it repeated the sequence and stayed in Protected View, with a "We've refreshed this document
  with the latest changes from others" bar.
- **Open:** whether that is the fake LOCK answer of the spike or the zone. Retest with the host in the Local intranet
  zone.

[MS-FSSHTTP]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/05fa7efd-48ed-48d5-8d85-77995e17cc81
[MS-WDVME]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wdvme/8cafdf55-ee5c-443e-bdb7-2cb2ab1fb2c3
[MS-WEBDAVE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-webdave/9fb175f3-f063-4389-a959-4457692e4b49
[MS-OFBA]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-ofba/30c7bbe9-b284-421f-b866-4e7ed4866027
[MS-FPSE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fpse/1bbae50a-7de9-45ea-830d-4e9caf8dfccd
[RFC7617]: https://www.rfc-editor.org/rfc/rfc7617

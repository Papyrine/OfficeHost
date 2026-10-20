# FSSHTTP spike (Phase 0)

The question: will desktop Word use [MS-FSSHTTP] (`cellstorage.svc`) against a host that is not SharePoint, and what
does it need to see first? This note has what the specs say, what Word did against the spike of the sample, how to run
the spike, how to capture Word's traffic with a real SharePoint, and what that capture has shown so far.

**Answer: yes.** Word uses MS-FSSHTTP against a host that is not SharePoint once the OPTIONS answer has two headers
that no published spec documents:
- `X-MSFSSHTTP: 1.3`;
- `X-MSWEBURL: <path of the site>`, the answer to Word's `X-MSGETWEBURL: t`.

Here is how that was found and confirmed:
1. **Spike without them:** across five runs, with four kinds of authentication and the SharePoint headers that the
   specs document, Word 16 (Microsoft 365, build 16.0.20326) opened the document over plain WebDAV.
2. **SharePoint Server, signed in:** a capture of the same Word against an on-premises SharePoint Server showed both
   headers on SharePoint's OPTIONS answer. Word POSTed to `/_vti_bin/cellstorage.svc` 50 ms later, with no FrontPage,
   `webs.asmx` or `_vti_inf.html` step.
3. **Spike with them (runs 6 and 7):** Word sent the spike the same first FSSHTTP request, `JoinCoauthoring` included.
   That held with Windows sign-in and with only a token in the path. The spike cannot answer Cell sub-requests yet,
   so Word then exited coauthoring and either failed the open (run 6) or fell back to a read-only GET (run 7).

So FSSHTTPB (the Cell sub-requests) is mandatory for an editable open, which is Phases 1 and 2.


## What the specs say

Section numbers below are those of the current Open Specifications. "Spec" marks what is documented, "trace" what
needs a capture.

- **No discovery in MS-FSSHTTP** (spec). [MS-FSSHTTP] 1.7, Versioning and Capability Negotiation, is "None".
  1.5 says the server endpoint is the site URL plus `/_vti_bin/cellstorage.svc`, or the whole document URL plus
  `/_vti_bin/cellstorage.svc`. It also says the server returns canonical URLs in `ResponseCollection/@WebUrl` (2.2.3.6)
  and `Response/@Url` (2.2.3.5). How a client decides that a server speaks FSSHTTP is not specified.
- **Transport** (spec, [MS-FSSHTTP] 2.1).
  - SOAP 1.1 over HTTP or HTTPS.
  - `SOAPAction: http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest`.
  - The server MUST answer in MTOM.
  - Faults are HTTP status codes or SOAP faults.
- **The first request** (spec, [MS-FSSHTTP] 4.1.1/4.1.2, "Successful File Open of a Coauthorable Document"). The
  open sends these sub-requests in one Request:
  1. Coauth `JoinCoauthoring` (with `SchemaLockID`, `ClientID`, `Timeout`, `AllowFallbackToExclusive`).
  2. SchemaLock `GetLock`, `DependsOn` the Coauth with `DependencyType="OnNotSupported"`.
  3. Three Cell sub-requests, `DependsOn` the SchemaLock with `OnExecute`. Two of them carry a `PartitionID` (content
     and metadata), one carries `GetFileProps`.
  4. `ServerTime` and `WhoAmI`.

  So the first FSSHTTP request already tries to join coauthoring; there is no `GetDocMetaInfo` in it. The spike answers
  the whole example except the Cell sub-requests (see below).
- **OPTIONS** (spec, [MS-WDVME] 3.2.5.1, 4.1.2; [MS-WEBDAVE] 2.2.1).
  - Required: `MS-Author-Via` with `DAV`.
  - Recommended: an `Allow` header that includes `GETLIB`, plus `DocumentManagementServer` and
    `Public-Extension: http://schemas.microsoft.com/repl-2`.
  - `X-MSDAVEXT: 1` advertises the WebDAV extended errors.
  - `MicrosoftSharePointTeamServices` is a SharePoint header without a spec section that defines it as a capability
    signal (trace).
- **FrontPage** (spec, [MS-FPSE]). `_vti_inf.html` and `_vti_bin/_vti_aut/author.dll` ("url to web url",
  "open service") are the old way to find the site of a document. Word sent none of them (see below).
- **Authentication** (spec, [MS-OFBA] 2.2.1, 2.2.2). A server that needs forms authentication answers 403 with
  `X-FORMS_BASED_AUTH_REQUIRED` (the login page) and `X-FORMS_BASED_AUTH_RETURN_URL` (the page that ends the login
  dialog).
- **Request headers Word sends that no spec found documents** (trace):
  - `X-MSGETWEBURL: t`
  - `X-MS-CookieUri-Requested: t`
  - `X-IDCRL_ACCEPTED: t`
  - `X-IDCRL_OPTIONS: force-auth-challenge`
  - `IgnoreCookieAuthentication: t`
  - `X-FeatureVersion: 1`
  - `X-MS-AuthDomainSupport: True`
  - `Accept-Auth: badger,Wlid1.1,Bearer,Basic,NTLM,Digest,Kerberos,Negotiate,Nego2`
  - an empty `Authorization: Bearer`

  `X-MSGETWEBURL` asks the server for the URL of the site in the OPTIONS answer. Without it, Word has no web URL, and
  that is the best candidate for why it stays on WebDAV.


## What Word did

All five runs used the same environment:
- **Client:** Windows 11 (10.0.26200), Word 16.0.20326 Click-to-Run.
- **Server:** `https://localhost:5001` with the ASP.NET Core development certificate, which is already trusted.
- **Office server cache:** `HKCU\Software\Microsoft\Office\16.0\Common\Internet\Server Cache` was absent, so empty.
- **Launch:** `Start-Process 'ms-word:ofe|u|<url>'`.

The logs of each run are in `src/WopiServer.Sample/spike-logs/run*` (ignored by git: they hold tokens).

| Run | URL / auth | OPTIONS answer | What Word did |
| --- | --- | --- | --- |
| 1 | `/t/{token}/Welcome.docx`, anonymous (token in the path) | `MicrosoftSharePointTeamServices`, `DAV: 1,2`, `MS-Author-Via: MS-FP/4.0,DAV`, `X-MSDAVEXT: 1`, `Allow`/`Public` with GETLIB | 7 × OPTIONS on the folder, HEAD, 3 × OPTIONS, LOCK (exclusive, `Timeout: Second-600`), GET, PROPFIND (Depth 0), UNLOCK, then OPTIONS every few seconds. It opened in **Protected View**. |
| 2 | as 1 | as 1, plus `DocumentManagementServer` and `Public-Extension` ([MS-WDVME] 4.1.2) | Same as run 1. |
| 3 | `/o/{token}/Welcome.docx`, [MS-OFBA] | 403 with `X-FORMS_BASED_AUTH_REQUIRED`/`_RETURN_URL`/`_DIALOG_SIZE` | 2 × OPTIONS, then a dialog: "Office has blocked this content because it uses a sign-in method that may be insecure", pointing to the Trust Center. It never opened the login page. |
| 4 | `/b/Welcome.docx`, Basic over HTTPS | 401 `WWW-Authenticate: Basic` | 2 × OPTIONS, 2 × HEAD, then a dialog: "Microsoft Office has blocked access to `https://localhost:5001/b/` because the source uses a sign-in method that may be unsecure". There was no credential prompt. |
| 5 | `/n/Welcome.docx`, Windows auth (Negotiate) | 401 `WWW-Authenticate: Negotiate`, then the run 2 headers | Word signed in silently (NTLM/Kerberos), then repeated run 1, in Protected View. After "Enable Editing" (pressed through UI Automation): LOCK, GET, PROPFIND, UNLOCK again, still in Protected View, and a "We've refreshed this document with the latest changes from others" bar. |
| 6 | `/n/Welcome.docx`, Windows auth | as 5, plus `X-MSFSSHTTP: 1.3`, and `X-MSWEBURL: /n` when asked (from the SharePoint capture below) | 3 × OPTIONS, then `POST /n/_vti_bin/cellstorage.svc/CellStorageService`: the first FSSHTTP request, as to SharePoint. The spike answered the Cell sub-requests `RequestNotSupported`, then: `sharedaccess.asmx` (404), 2 × `ExitCoauthoring`, and "Sorry, we couldn't open 'https://localhost:5001/n/Welcome.docx'". No HEAD, GET or LOCK. |
| 7 | `/t/{token}/Welcome.docx`, anonymous (token in the path) | as 6, with `X-MSWEBURL: /t/{token}` | 2 × OPTIONS, then the same FSSHTTP open, `JoinCoauthoring` included, with no HTTP sign-in. After the Cell refusal: `sharedaccess.asmx` (404), 2 × `ExitCoauthoring`, then a fallback: GET of the document, and at the host root `_vti_inf.html`, the FrontPage RPC, HEAD, `webs.asmx` `WebUrlFromPageUrl`, `workflow.asmx` `GetWorkflowDataForItem`. It opened read-only in Protected View. |

Across all runs:
- **`cellstorage.svc` was never reached**, so neither `JoinCoauthoring` nor any other FSSHTTP sub-request was ever
  sent.
- **No FrontPage, SOAP or REST probe:** no `_vti_inf.html`, `author.dll`, `shtml.dll`, `webs.asmx`, `lists.asmx` or
  `/_api`, even with `MS-Author-Via: MS-FP/4.0` in every OPTIONS answer.
- **The OPTIONS requests come in two kinds:**
  - most carry `X-MSGETWEBURL: t`;
  - the first of each authenticated run carries `X-IDCRL_OPTIONS: force-auth-challenge` and
    `IgnoreCookieAuthentication: t` instead, which asks for an authentication challenge.
- **Two user agents:** `Microsoft Office Word 2014 (16.0.20326) Windows NT 10.0` for most requests, and plain
  `Microsoft Office Word` for some of the OPTIONS.
- **The LOCK body** names the Windows identity of the user: `<D:owner><D:href>AzureAD\…</D:href></D:owner>`.
- **The PROPFIND asks** for `creationdate`, `getlastmodified`, `getcontentlength` and `Office:modifiedby`
  (`urn:schemas-microsoft-com:office:office`).

Three conclusions follow:
1. **Default sign-in policy.** Basic and MS-OFBA are blocked by the default Office policy for "insecure sign-in
   methods". What remains by default is Bearer (Microsoft Entra ID, which SharePoint Online uses) and Windows auth.
   Windows auth works, but only for domain or Entra-joined clients against an intranet host.
2. **Authentication is not the trigger.** Anonymous, token in the path and Windows auth all led to the same WebDAV
   open.
3. **Protected View confounds the test.** `https://localhost` is in the Internet zone, so Word opens read-only and
   releases its WebDAV lock straight away. A SharePoint site in the Trusted sites or Local intranet zone would not hit
   this. Run 5 is the closest to that, with Windows auth and "Enable Editing", and still nothing reached `_vti_bin`.


## Running the spike

The spike is in the sample only (`src/WopiServer.Sample/FsshttpSpike.cs`). It is off unless `Spike:Fsshttp` is true,
and the `spike` launch profile turns it on, with Windows auth, on `https://localhost:5001`.

1. **Trust the HTTPS development certificate.** It goes into the CurrentUser root store, which Word uses.

   ```bash
   dotnet dev-certs https --trust
   ```
2. **Start the sample with the spike profile:**

   ```bash
   dotnet run --project src/WopiServer.Sample --launch-profile spike
   ```
3. **Open `https://localhost:5001/spike`.** It has an `ms-word:ofe|u|…` link for each site:

   | Site | Auth |
   | --- | --- |
   | `/t/{token}` | token in the path |
   | `/o/{token}` | MS-OFBA, signing in from the token without a prompt |
   | `/b` | Basic, with the token as the password |
   | `/n` | Windows auth |
   | `/` | anonymous |

   Click a link, or run it from PowerShell:

   ```bash
   pwsh -c "Start-Process 'ms-word:ofe|u|https://localhost:5001/n/Welcome.docx'"
   ```
4. **Read the logs.** Every request and response is written in full to `src/WopiServer.Sample/spike-logs`, one file
   each, numbered. Binary and multipart bodies are written as a hex dump, with the parts of an MTOM body decoded. The
   console also logs one line per request.
5. **Try headers from a capture without rebuilding.** Add them to every OPTIONS answer through configuration, as
   environment variables in the launch profile or on the command line, for example:

   ```bash
   dotnet run --project src/WopiServer.Sample --launch-profile spike -- --Spike:OptionsHeaders:X-Some-Header=value
   ```
6. **Clear Office's per-server cache between runs** if it exists: the key
   `HKCU\Software\Microsoft\Office\16.0\Common\Internet\Server Cache`. Close Word normally. Killing it makes the next
   start offer Safe Mode.

What the spike answers:
- **OPTIONS:** as in run 6: the headers of run 2, `X-MSFSSHTTP: 1.3`, and `X-MSWEBURL` (the site path) when asked.
- **`_vti_inf.html`:** [MS-FPSE].
- **FrontPage RPC:** "server version", "url to web url" and "open service".
- **`webs.asmx`:** `WebUrlFromPageUrl`.
- **WebDAV:** GET, HEAD and PUT of the document, PROPFIND, and a fake exclusive LOCK/UNLOCK.
- **`cellstorage.svc`** (at the site or at the document URL), always in MTOM:
  - **Success:** `WhoAmI`, `ServerTime` and `GetDocMetaInfo`.
  - **Stub success:** Coauth (always `Alone`), SchemaLock, ExclusiveLock and EditorsTable.
  - **`DependentOnlyOnNotSupportedRequestGetSupported`:** an `OnNotSupported` dependent whose dependency was
    supported, as in [MS-FSSHTTP] 4.1.2.
  - **`RequestNotSupported`:** Cell, because there is no FSSHTTPB codec yet.

`FsshttpSpikeTests` posts the request of [MS-FSSHTTP] 4.1.1 to it and checks these answers.


## Capturing Word with a real SharePoint

[src/fsshttp-capture.ps1](../src/fsshttp-capture.ps1) captures Word's HTTP traffic without a proxy and without
trusting a new root certificate.
- **How it records:** it starts an ETW trace of WebIO (the engine of WinHTTP, which Word uses for documents), WinINet
  and WinHTTP. It opens the document in Word, stops the trace when you press Enter, and writes the requests and
  responses to that host as text: headers, and the bytes sent and received. ETW logs those bytes before TLS, so they
  are in plain text.
- **What it needs:** membership of Performance Log Users (or an elevated prompt).
- **What it redacts:** Authorization, cookie and ticket values in the text.
- **Keep the trace private:** the `.etl` file holds credentials.
- **No bodies from WinINet:** the WinINet-Capture provider, which would log WinINet bodies, needs an elevated prompt.
  Word only used WinINet for its own sign-in configuration.

Run it with the document URL:

```bash
pwsh src/fsshttp-capture.ps1 -Url 'https://sharepoint.example/Shared%20Documents/test.docx'
```

### SharePoint Server, first capture

The capture used the "Edit brief" button of an application on SharePoint Server (`MicrosoftSharePointTeamServices:
16.0.0.5522`, IIS 10, forms-based sign-in). That button calls `window.open('ms-word:ofe|u|https://<host>/<library>/<folder>/<file>.docx')`.
- **Word's first request** to the host was the same OPTIONS, on the folder, that it sends the spike:
  `Authorization: Bearer` (empty), `X-MSGETWEBURL: t`, `X-MS-CookieUri-Requested: t`, `X-IDCRL_ACCEPTED: t`,
  `Accept-Auth: …`. The user agent was plain `Microsoft Office Word 2014`.
- **SharePoint's answer** was `403 FORBIDDEN` with:

  ```
  X-Forms_Based_Auth_Required: https://<host>/_login/default.aspx?ReturnUrl=/_layouts/15/error.aspx
  X-Forms_Based_Auth_Return_Url: https://<host>/_layouts/15/error.aspx
  X-MSDAVEXT_Error: 917656; Access+denied.+Before+opening+files+in+this+location%2c+you+must+first+browse+to+the+web+site+and+select+the+option+to+login+automatically.
  X-MS-InvokeApp: 1; RequireReadOnly
  MicrosoftSharePointTeamServices: 16.0.0.5522
  X-SharePointHealthScore: 0
  SPRequestGuid: …
  ```
- **Word's reaction:** it showed "Office has blocked this content because it uses a sign-in method that may be
  insecure", exactly as with the MS-OFBA site of the spike (run 3). It sent nothing more until the dialog was
  dismissed.

What this capture shows:
1. **The OFBA block is Office's default policy.** It is not something the spike got wrong: the policy blocks
   SharePoint's own forms-based sign-in too. It is File > Options > Trust Center > Trust Center Settings >
   Form-based Sign-in (Group Policy: "Control how Office handles form-based sign-in prompts"). Its default blocks all
   prompts, and it can allow specific hosts.
2. **An unauthenticated OPTIONS learns nothing.** SharePoint answers it with the sign-in challenge alone, without
   `DAV`, `MS-Author-Via` or any web URL. So the answer to `X-MSGETWEBURL` comes only after sign-in. The next capture
   has to get past sign-in: allow forms-based sign-in for this host in the Trust Center, or use a SharePoint with
   Windows or Entra ID sign-in.

### SharePoint Server, second capture (forms-based sign-in allowed)

With Form-based Sign-in set to "Ask me what to do for each host", and the host allowed, Word got further:
1. **OPTIONS:** the same OPTIONS and 403 as before.
2. **HEAD of the document:** 403 with the same sign-in headers.
3. **REST v2:** `GET /_api/v2.0/shares/u!<base64url of the document URL>/driveitem/?$select=viewpoint`, answered
   `401 NTLM`. This is Word asking SharePoint's REST v2 API (OneDrive shape) about the item. So Word treats the host
   as SharePoint even before sign-in.
4. **The login page:** Word opened it in its sign-in dialog. That dialog is the IE engine, going through WinINet inside
   Word.
5. **Sign-in failed both ways** on this PC, which is not on the SharePoint domain:
   - SharePoint's "Windows Authentication" went to NTLM;
   - "ADFS" went to the ADFS Windows-integrated endpoint, because ADFS sends IE there.

   Neither asked for a password; the dialog showed "Can't reach this page" (details in
   [fsshttp-notes.md](fsshttp-notes.md)).

So no FSSHTTP traffic yet. Getting it needs Word to sign in, which means one of these:
- a PC on the SharePoint domain (Windows sign-in is then silent);
- domain credentials saved for the host, with the host in the Local intranet zone;
- an ADFS that gives IE the password form;
- SharePoint Online, where Word signs in with Entra ID.

### SharePoint Server, third capture (signed in)

Signed in from an earlier session (a `FedAuth` cookie), Word opened the document in edit mode. In order:
1. **3 × OPTIONS on the folder,** each answered 200 with the WebDAV and FrontPage headers of run 2 and `X-MSFSSHTTP: 1.3`.
   The two that carried `X-MSGETWEBURL: t` also got `X-MSWEBURL: /` (the site is the root of the host).
2. **`POST /_vti_bin/cellstorage.svc/CellStorageService`:** the URL is the web URL plus the service path, MS-FSSHTTP 1.5.
   - **HTTP headers:** `SOAPAction: "http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest"`,
     plus `Application: Win32-Word`, `FileExtension: .docx`, `Scenario: OpenFile_Open`, `ScenarioType: AUO`, and
     `X-RequestStats: Operation=Open;…`.
   - **Body:** MTOM (`multipart/related; type="application/xop+xml"; start="<http://tempuri.org/xml/0>"`), with one XML
     part.
   - **`RequestVersion`:** `Version="2" MinorVersion="2"`.
   - **`Request` attributes:** `Url`, `UseResourceID="true"`, `ResourceID` (from an earlier open), `UserAgent`,
     `UserAgentClient="msword"`, `UserAgentPlatform="win"`, `Build`, `RequestToken="1"`, `MetaData="1031"`.
   - **Sub-requests, in order:**
     1. Coauth `JoinCoauthoring` (`SchemaLockID="{29358EC1-E813-4793-8E70-ED0344E7B73C}"`, `Timeout="2580"`,
        `AllowFallbackToExclusive="true"`);
     2. SchemaLock `GetLock`, `OnNotSupported` the join;
     3. three Cell sub-requests, `OnExecute` the SchemaLock: the metadata partition `7808F4DD-…`, the content partition
        `383ADC0B-…`, and one with `GetFileProps="true"`. Their FSSHTTPB is inline base64, not an XOP part;
     4. three EditorsTable `UpdateEditorMetadata`, `OnSuccess` the join, with keys `SupportsRename` (`false`),
        `Capabilities` (`{"CTID":null,"ENV":0,"SF":0}`) and `Platform` (`1`);
     5. `WhoAmI`;
     6. `GetDocMetaInfo` (`OnExecute` WhoAmI);
     7. `GetVersions`;
     8. `ServerTime`;
     9. a last Cell sub-request of 27 bytes.
3. **The response:** `multipart/related` MTOM, chunked.
   - **Envelope:** `ResponseVersion Version="2" MinorVersion="3"`, and `ResponseCollection WebUrl="https://<host>"`
     without `WebUrlIsEncoded`.
   - **`Response`:** its `Url` is unencoded, spaces included, and it echoes `ResourceID`.
   - **Sub-responses:** the join answered `LockType="SchemaLock" CoauthStatus="Alone"` with `TransitionID` equal to the
     `ResourceID`. The SchemaLock got `DependentOnlyOnNotSupportedRequestGetSupported`, `HResult="2147500037"`, as in
     MS-FSSHTTP 4.1.2. `WhoAmI` returned a claims login (`i:0#.w|<domain>\<user>`) and an empty `UserSIPAddress`.
   - **Cell results** carry `ServerCorrelationId`. They come either inline as base64 or as an XOP part
     (`<xop:Include href="cid:http://tempuri.org/1/…"/>`, `Content-Transfer-Encoding: binary`). ETW logs only the start
     of large receive buffers, so the rest of this response was cut.
4. **`POST /_vti_bin/sharedaccess.asmx`, `IsOnlyClient` ([MS-SHDACCWS]):**
   - the body is `<id>` = the `ResourceID`, and the answer `true`;
   - the headers are `Scenario: PollingSubscriptionNotConnected_AmIAlone` and `X-RequestStats: Operation=AmIAlone`;
   - it repeats every 8 seconds for a minute, then every 30 seconds.
5. **`OPTIONS /`, then REST** `POST /_api/web/GetFileByUrl(@v)/ListItemAllFields/GetSharingInformation`. That got
   403 (no digest), then, with the `X-RequestDigest` it returned, 404. These are not needed for the open.
6. **The save:** you typed a word and saved. Word showed "REQUIRED PROPERTIES: To save to the server, correct the invalid
   or missing required properties", and sent no further cellstorage request. The library has required columns, and
   Word checks them before uploading. So PutChanges is not in this capture.

### What later captures should do

The open is captured. What SharePoint does after it is not:
- the save (PutChanges);
- the refresh cadence (`RefreshCoauthoring`, EditorsTable);
- two users coauthoring.

These need:
1. **A library without required columns,** so that Word uploads the save.
2. **Run `src/fsshttp-capture.ps1`** with the document URL, signed in. Then:
   1. Wait for edit mode.
   2. Type a word and save.
   3. Leave Word open for 15 minutes (`Timeout="2580"` suggests minutes, not seconds, between refreshes).
3. **A second user** on the same document at the same time, editing another paragraph and saving.

What ETW gives here: WebIO logs only the start of large receive buffers, so long SharePoint responses come out cut.
The spike logs everything Word sends it in full. So from Phase 1 on, Word against the spike is the main source for
requests, and SharePoint captures are the reference for responses.

What to look for: the sizes and the FSSHTTPB of the Cell sub-requests at save (`BinaryDataSize`, `Coalesce`, XOP
parts), `RefreshCoauthoring` and its interval, EditorsTable `RefreshSession`, and with two users
`CoauthStatus="Coauthoring"` and a PutChanges that fails coherency and is retried.


## Go / no-go

**Go.** Word sends its FSSHTTP open, `JoinCoauthoring` included, to a host that is not SharePoint, as soon as OPTIONS
answers with `X-MSFSSHTTP` and `X-MSWEBURL` (run 6). That leaves the Cell sub-requests (FSSHTTPB) as the
missing piece, which is exactly what Phases 1 and 2 build.

What the later phases depend on, and what is still open:
1. **Cell is mandatory.** With the Cell sub-requests refused, Word exits coauthoring. It then either fails the open
   (run 6) or falls back to a read-only GET (run 7), but never edits. Phase 1 has real input: the base64 FSSHTTPB of
   Word's Cell sub-requests is in the logs of runs 6 and 7 (no document content, they are queries), next to the
   vectors of [MS-FSSHTTP] 4.1.
2. **`sharedaccess.asmx` `IsOnlyClient` ([MS-SHDACCWS])** is polled from the open on. It is a small SOAP service, and
   Phase 4 has to add it to `MapFsshttp()`.
3. **Sign-in.**
   - Office blocks Basic and forms-based sign-in by default (for SharePoint too), and the forms-based option is being
     retired.
   - Windows sign-in works (run 6).
   - So does a token in the path with no HTTP sign-in (run 7). That maps directly onto `IWopiTokens`, as the token of
     a WOPI launch does, and is the natural default for Phase 4 and the desktop launch URL of `WopiLauncher`.
   - Entra ID (Bearer) is untested.
4. **Protected View.** It no longer gets in the way of the open: run 6 went to FSSHTTP from the Internet zone. Whether
   Word then edits, or stays read-only, shows once the open succeeds, in Phase 2.
5. **Save and coauthoring traffic** (PutChanges, refresh intervals, two users) are not captured yet. See "What later
   captures should do".


## Licence note

Nothing has been copied. OfficeDev/Interop-TestSuites (MIT) has FSSHTTPB structures that may be read for understanding
in Phase 1. If any of it is adapted, the MIT notice goes with it and the PR says so.

[MS-FSSHTTP]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/05fa7efd-48ed-48d5-8d85-77995e17cc81
[MS-FSSHTTPB]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/f59fc37d-2232-4b14-baac-25f98e9e7b5a
[MS-WDVME]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wdvme/8cafdf55-ee5c-443e-bdb7-2cb2ab1fb2c3
[MS-WEBDAVE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-webdave/9fb175f3-f063-4389-a959-4457692e4b49
[MS-OFBA]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-ofba/30c7bbe9-b284-421f-b866-4e7ed4866027
[MS-FPSE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fpse/1bbae50a-7de9-45ea-830d-4e9caf8dfccd

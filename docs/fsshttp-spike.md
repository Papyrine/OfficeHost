# FSSHTTP spike (Phase 0)

The question: will desktop Word use [MS-FSSHTTP] (`cellstorage.svc`) against a host that is not SharePoint, and what
does it need to see first? This note has what the specs say, what Word did against the spike of the sample, how to run
the spike, how to capture Word's traffic with a real SharePoint, and what that capture has shown so far.

**Answer so far: not yet.**
- **Against the spike:** across five runs, with four kinds of authentication and the SharePoint headers that the specs
  document, Word 16 (Microsoft 365, build 16.0.20326) opened the document over plain WebDAV. It never requested
  `_vti_inf.html`, the FrontPage RPC, `webs.asmx`, `/_api` or `cellstorage.svc`.
- **Against SharePoint Server:** a capture of the same Word on an on-premises SharePoint Server (forms-based sign-in)
  stopped at the first OPTIONS. Office blocks forms-based sign-in by default, for SharePoint too.

The specs do not document the step that makes Word switch to FSSHTTP, so that step still has to come from a capture of
Word getting past sign-in.


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
- **OPTIONS:** as in run 2.
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

### What the next capture should do

1. **Close every Office app.** Then delete `HKCU\Software\Microsoft\Office\16.0\Common\Internet\Server Cache` (if it
   exists), so that Word rediscovers the server.
2. **Run `src/fsshttp-capture.ps1`** with the document URL. Sign in when Word asks.
3. **Generate the coauthoring traffic:**
   1. Wait until Word shows the document in edit mode.
   2. Type a word and save.
   3. Wait two minutes (refresh timers).
   4. Close Word, then press Enter.
4. **Repeat with a second user,** for the coauthoring half: have a second user open the same document at the same
   time, in a second Windows session or on a second machine, edit another paragraph and save.

What to look for, in order of importance:
- **The OPTIONS answer to `X-MSGETWEBURL: t`:** every response header, and especially any that carries the site URL
  or a cookie URI (`X-MS-CookieUri-Requested`). Then put those headers into
  `Spike:OptionsHeaders` and rerun run 5.
- **The first request after the OPTIONS:**
  - is it `cellstorage.svc` at the site or at the document URL?
  - is it preceded by HEAD, PROPFIND or GET of the document?
  - does any request go to `_vti_bin`, `/_api` or `/_layouts`?
- **The auth exchange:**
  - with forms-based sign-in: the login dialog's requests, and the cookie (`FedAuth`) set after it, and on which path;
  - with Entra ID: the 401 with `WWW-Authenticate: Bearer …` (`client_id`, `trusted_issuers`, `authorization_uri`)
    and `X-IDCRL_AUTH_PARAMS_V1`;
  - whether the OPTIONS answer changes once Word is signed in.
- **The first `ExecuteCellStorageRequest`:**
  - the `RequestVersion` (`Version`, `MinorVersion`);
  - the `Request` attributes (`UserAgent`, `UserAgentClient`, `UserAgentPlatform`, `Build`, `MetaData`,
    `ShouldReturnDisambiguatedFileName`, …);
  - the list of sub-requests with their `DependsOn` and `DependencyType`;
  - the `SchemaLockID` and `ClientID` values;
  - `CoauthRequestType` versus `ExclusiveLockRequestType`.
- **The responses:**
  - `ResponseVersion`, `WebUrl` and `Url`, `HealthScore`;
  - `CoauthStatus`, `TransitionID`, and the SharePoint headers on the response (`SPRequestGuid`,
    `X-SharePointHealthScore`, `MicrosoftSharePointTeamServices`).
- **The refresh cadence:** how often `RefreshCoauthoring` and EditorsTable `RefreshSession` come, and with which
  `Timeout`.
- **The sizes of the Cell sub-requests:** QueryChanges at open, PutChanges at save (`BinaryDataSize`, the MTOM part
  sizes, `Coalesce`).
- **With two users:** `CoauthStatus="Coauthoring"`, the EditorsTable entries, and a PutChanges that fails coherency
  and is retried.


## Go / no-go

**No-go for Phase 1 until the capture is replayed.** Phases 1 to 5 are weeks of work, and nothing so far shows that
Word will call `cellstorage.svc` on a host that is not SharePoint. The evidence above says the trigger is neither the
documented OPTIONS headers nor authentication.

The remaining unknowns, and how each gets settled:
1. **What answers `X-MSGETWEBURL`** (and `X-MS-CookieUri-Requested`).
   - Settled by: a capture of Word signed in to SharePoint (the first SharePoint Server capture stopped at sign-in),
     then a rerun of run 5 with those headers.
   - If it works: Word reaches `cellstorage.svc`, and Phase 0 reruns to see the first request. Go.
2. **Which sign-in a host can use.** Office blocks Basic and forms-based sign-in by default, for SharePoint as well.
   Users or administrators can allow forms-based sign-in per host, but the option is being retired. The defaults
   leave Windows sign-in and Entra ID (Bearer).
   - Settled by: the signed-in capture, which shows whether Word gates FSSHTTP on how it signed in, then a spike that
     answers Bearer for an Entra app registered in your tenant.
   - Bearer works but nothing else does: go only for hosts whose users sign in with Entra ID or Windows, and add
     OAuth to the plan.
3. **Protected View** on a host in the Internet zone.
   - Settled by: adding `https://localhost` to Local intranet or Trusted sites. It is a security setting, so it is your
     call, and I have not changed it.
   - What it changes: the spike is retested in edit mode. Production hosts would need the same zone assignment, by
     policy.
4. **If 1 to 3 still give WebDAV only:** no-go for an FSSHTTP host. The documented route to coauthoring with Microsoft
   clients on third-party storage is then the Cloud Storage Partner Program "WOPI extensions for coauthoring" (CSPP
   Plus). It needs Microsoft onboarding, and the coauthoring runs through Microsoft's service.

Phase 1 can start in parallel only if you accept the risk. The FSSHTTPB codec is self-contained and tested against the
spec vectors (the base64 bodies of [MS-FSSHTTP] 4.1 and the [MS-FSSHTTPB] examples). If Word never reaches
`cellstorage.svc`, though, it has no user.


## Licence note

Nothing has been copied. OfficeDev/Interop-TestSuites (MIT) has FSSHTTPB structures that may be read for understanding
in Phase 1. If any of it is adapted, the MIT notice goes with it and the PR says so.

[MS-FSSHTTP]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/05fa7efd-48ed-48d5-8d85-77995e17cc81
[MS-FSSHTTPB]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/f59fc37d-2232-4b14-baac-25f98e9e7b5a
[MS-WDVME]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wdvme/8cafdf55-ee5c-443e-bdb7-2cb2ab1fb2c3
[MS-WEBDAVE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-webdave/9fb175f3-f063-4389-a959-4457692e4b49
[MS-OFBA]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-ofba/30c7bbe9-b284-421f-b866-4e7ed4866027
[MS-FPSE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fpse/1bbae50a-7de9-45ea-830d-4e9caf8dfccd

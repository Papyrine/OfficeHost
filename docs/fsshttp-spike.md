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

The logs of each run are in `src/OfficeHost.Sample/spike-logs/run*` (ignored by git: they hold tokens).

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


## The spike

The spike was throwaway code in the sample (`src/OfficeHost.Sample/FsshttpSpike.cs`, the `spike` launch profile):
SharePoint-like discovery answers, FrontPage RPC, WebDAV, and a cell storage service that answered WhoAmI, ServerTime
and GetDocMetaInfo, stubbed the locks and refused Cell, on sites with a token in the path, MS-OFBA, Basic, Windows
auth, or none. It was removed once the library served desktop Word (`MapFsshttp`); it is in the history of the
`fsshttp` branch up to commit 4f7748f. What is left of it is the request log of the sample's `desktop` launch profile
(`RequestLog.cs`, to `src/OfficeHost.Sample/request-logs`), and the check by hand with Word is
[fsshttp-checklist.md](fsshttp-checklist.md).


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

### SharePoint Server, fourth capture (close, reopen, save attempt)

1. **Close:** closing the document sent one cellstorage request, Coauth `ExitCoauthoring` with SchemaLock
   `ReleaseLock` `OnNotSupported` it.
2. **Reopen:** this sent the same open as the third capture.
3. **Refresh:** in the five minutes Word then stayed open, it sent no `RefreshCoauthoring` and no EditorsTable refresh.
   With `Timeout="2580"`, the refresh is later than that.
4. **Polling:** the `IsOnlyClient` polling went from 30 seconds back to 8 seconds after the edit.
5. **The save** was set up through Word's COM automation: the edit, the library's required columns (the server's own
   values, plus the content type's required `Policy` and `SubPolicy`), then `Save()`, then the "Retry Save" button.
   Each attempt failed in Word with "Required Properties" and sent nothing to the server. Word's property pane said "No
   SharePoint properties found", so Word did not get the library's property schema for this document. That is a matter
   of this library's setup, not of FSSHTTP. The unsaved changes were discarded.

### SharePoint Server, fifth capture (two saves)

This used another brief, whose property schema Word did load: it opened, in Compatibility Mode, with the same
FSSHTTP open. The edits and saves went through Word's COM automation (`Range.InsertAfter`, `Document.Save`), about a
minute apart. `Save()` returned within a second, because Word saves to its cache and uploads in the background.

Each save is:
1. **The upload: `POST /_vti_bin/cellstorage.svc/CellStorageService`.** It has `Scenario: AutoSaveFile_Save` and
   18,506 bytes of body for the first save, 17,612 for the second. The body is:

   ```xml
   <Request Url="…" UseResourceID="true" … MetaData="7" ResourceID="…">
     <SubRequest Type="Cell" SubRequestToken="1">
       <SubRequestData Coalesce="true" CoauthVersioning="true" GetFileProps="true"
           BypassLockID="{29358EC1-E813-4793-8E70-ED0344E7B73C}" SchemaLockID="{29358EC1-E813-4793-8E70-ED0344E7B73C}"
           BinaryDataSize="17026">
         <xop:Include href="cid:http://tempuri.org/0" xmlns:xop="http://www.w3.org/2004/08/xop/include"/>
       </SubRequestData>
     </SubRequest>
   </Request>
   ```

   - **One sub-request:** a single Cell. There is no Coauth, refresh or lock sub-request with it.
   - **The lock:** `BypassLockID` is the schema lock of the coauthoring session (MS-FSSHTTP 2.3.3.1).
   - **`MetaData`:** `7`, where the open had `1031`.
   - **The FSSHTTPB (PutChanges):** a binary XOP part (`Content-Transfer-Encoding: binary`), not inline base64 as at
     open.
   - **The response:** `200`, MTOM, `MinorVersion="3"`. The rest was cut by ETW.
2. **The metadata: `POST cellstorage.svc`, `Scenario: GetMetadata_EnsureDocMetaInfo`.** It is a `GetDocMetaInfo`
   sub-request alone, `MetaData="1031"`. SharePoint's `DocProps` are the `vti_*` properties of the file, mixed with the
   library's columns:
   - **The lock:** `vti_sourcecontrollockid` (the `SchemaLockID`), `vti_sourcecontrollocktype` (`1`),
     `vti_sourcecontrollockexpires`, `vti_sourcecontrolcheckedoutby`, `vti_sourcecontrolmultiuserchkoutby`,
     `vti_sourcecontroltimecheckedout`.
   - **Versions:** `vti_contenttag` (`{ResourceID},<version>,<n>`), `vti_contentversionisdirty`,
     `vti_TimeOfLastCobaltVersionCreation`, `vti_internalversion`, `vti_nextbsn`, `vti_metadatanextbsn`,
     `vti_nextbsnallocated`.
   - **The file:** `vti_filesize`, `vti_timecreated`, `vti_timelastmodified`, `vti_author`, `vti_modifiedby` (claims
     logins), `vti_title`, `vti_replid` (`rid:{ResourceID}`), `vti_parentid`.
   - **The store:** `vti_docstoretype`, `vti_docstoreversion`, `vti_streamschema`, `vti_parserversion`,
     `vti_level`, `vti_canmaybeedit`, `vti_ZeroByteFileIsOverwriteable`, `ContentTypeId`.
   - **The library's columns.**
3. **REST:** the REST `GetSharingInformation` call again.

No `RefreshCoauthoring` came in the four minutes of this capture.

### SharePoint Server, sixth capture (two users coauthoring)

Word on this machine (user 1) had the brief open for 27 minutes. A second user, on another machine, opened it, edited
and saved twice. User 1's edits and saves went through COM. The lab connection dropped for three minutes (DNS "No such
host is known"), so user 1's first save waited in Word's cache and uploaded on reconnect. Word's `CoAuthoring`
object showed 2 authors, then "pending updates", then the merge. The trace saw only user 1's side: 166 cellstorage
requests.

1. **While alone,** Word polled `sharedaccess.asmx` `IsOnlyClient` (8 seconds, then 30), and nothing else.
2. **Back online, with a second user in** (14:48):
   - `GetDocMetaInfo`.
   - **`Scenario: GetChangesFile_MarkCoauthTransitionCompleteAndDownload`:** Coauth `MarkTransitionComplete`, plus the
     three Cell queries of the open (`GetFileProps`, content, metadata). The answer had the Etag at version 15, with the
     second user's saves.
   - **`Scenario: Coauthoring_RequestUploadOnSecondaryMetadataPartition`:** a Cell PutChanges of 1,302 bytes alone.
   - **The save** (`AutoSaveFile_Save`): PutChanges of 20,028 bytes. The Etag went to version 16, with no coherency
     failure, because Word had pulled the remote changes first.
3. **Then the poll.** `IsOnlyClient` stopped, and Word sent the same request **every 8 seconds** for the rest of the
   session (159 times), `Scenario: PollingSubscriptionNotConnected_TransitionToCoauthInHostMode`. It holds:
   - three Cell queries (`GetFileProps="true"`, then two more; 539, 508 and 388 bytes of FSSHTTPB);
   - Coauth `GetCoauthoringStatus`, answered `CoauthStatus="Coauthoring"` every time.

   The Cell queries are QueryChanges with Word's knowledge. When the second user saved (Etag version 17, 14:51:29),
   the next answer was larger, with their changes. From then on Word's queries changed size (539 → 537 bytes), as its
   knowledge did. Word kept the changes as "pending updates" and merged them into the document on its next save
   (14:52:02, PutChanges of 17,396 bytes, Etag version 18).
4. **Refresh:** no `RefreshCoauthoring`, no EditorsTable refresh, and no SchemaLock refresh came in the 27 minutes
   since the open. With `Timeout="2580"` (43 minutes) the refresh is later. `GetCoauthoringStatus` every 8 seconds is
   what tells Word who else is in.

### What later captures should do

Open, save, polling and the merge of a remote change are captured. Still missing:
- **The refresh:** leave Word open for 45 minutes to see `RefreshCoauthoring` (or what replaces it).
- **A conflicting save:** both users edit the same paragraph and save. That should give a coherency failure, or the
  "same paragraph" lock that Word shows on SharePoint.
- **The second user's side:** run `src/fsshttp-capture.ps1` on their machine too.
- **Full response bodies:** from Phase 2, the spike gets Word's saves itself.

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
   - So does a token in the path with no HTTP sign-in (run 7). That maps directly onto `IOfficeTokens`, as the token of
     a WOPI launch does, and is the natural default for Phase 4 and the desktop launch URL of `OfficeLauncher`.
   - Entra ID (Bearer) is untested.
4. **Protected View.** It no longer gets in the way of the open: run 6 went to FSSHTTP from the Internet zone. Whether
   Word then edits, or stays read-only, shows once the open succeeds, in Phase 2.
5. **The save is captured** (fifth capture): one Cell sub-request with `Coalesce`, `CoauthVersioning`,
   `GetFileProps` and `BypassLockID`, then a `GetDocMetaInfo`.
6. **Coauthoring is captured** (sixth capture). It is driven by a poll every 8 seconds, Cell QueryChanges plus Coauth
   `GetCoauthoringStatus`, so the cost of the server is that poll: Phase 3's store has to answer QueryChanges against
   a client's knowledge cheaply. Also needed:
   - `MarkTransitionComplete` when a second user arrives;
   - a separate PutChanges of the metadata partition.

   The refresh (beyond 27 minutes) and a conflicting save are not captured yet.


## Licence note

Nothing has been copied. OfficeDev/Interop-TestSuites (MIT) has FSSHTTPB structures that may be read for understanding
in Phase 1. If any of it is adapted, the MIT notice goes with it and the PR says so.

[MS-FSSHTTP]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/05fa7efd-48ed-48d5-8d85-77995e17cc81
[MS-FSSHTTPB]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/f59fc37d-2232-4b14-baac-25f98e9e7b5a
[MS-WDVME]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wdvme/8cafdf55-ee5c-443e-bdb7-2cb2ab1fb2c3
[MS-WEBDAVE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-webdave/9fb175f3-f063-4389-a959-4457692e4b49
[MS-OFBA]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-ofba/30c7bbe9-b284-421f-b866-4e7ed4866027
[MS-FPSE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fpse/1bbae50a-7de9-45ea-830d-4e9caf8dfccd

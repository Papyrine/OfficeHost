# FSSHTTP: where Word differs from the specs

Each entry gives the spec section, what Word did, and the evidence: a run of the spike (docs/fsshttp-spike.md), whose logs are in `src/OfficeHost.Sample/spike-logs`, or a run of the library through the `desktop` launch profile of the sample, whose logs are in `src/OfficeHost.Sample/request-logs` (those of the earliest runs in `spike-logs` too). The client is Word 16.0.20326 (Microsoft 365) on Windows 11 10.0.26200.

Several editors, the locks and the coherency of a save (coauthoring, Excel's exclusive lock) are in [fsshttp-concurrency.md](fsshttp-concurrency.md).


## Discovery

### No FrontPage RPC despite `MS-Author-Via: MS-FP/4.0`

- **Spec:** [MS-WDVME] 3.2.5.1 and 4.1.2 have a FrontPage-compliant server advertise `MS-Author-Via: MS-FP/4.0,DAV`. [MS-FPSE] defines `_vti_inf.html` and `author.dll` for clients that use it.
- **Word:** in runs 1, 2 and 5 every OPTIONS answer had `MS-Author-Via: MS-FP/4.0,DAV`, `DocumentManagementServer` and `MicrosoftSharePointTeamServices`. Word sent no FrontPage request at all (`_vti_inf.html`, `shtml.dll`, `author.dll`), and no `webs.asmx` either. It used WebDAV only: OPTIONS, HEAD, LOCK, GET, PROPFIND, UNLOCK.

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

  Some OPTIONS used the user agent `Microsoft Office Word` alone. In each run, one OPTIONS replaced `X-MSGETWEBURL` with `X-IDCRL_OPTIONS: force-auth-challenge` and `IgnoreCookieAuthentication: t`.
- **Open:** the server's answer to `X-MSGETWEBURL` and `X-MS-CookieUri-Requested`. It needs a SharePoint Online capture.

### FSSHTTP is negotiated on OPTIONS: X-MSFSSHTTP and X-MSWEBURL

- **Spec:** [MS-FSSHTTP] 1.7 (Versioning and Capability Negotiation) is "None", and 1.5 only says where the endpoint is. Neither header is in [MS-FSSHTTP], [MS-WDVME], [MS-WEBDAVE] or [MS-OFBA].
- **SharePoint Server 16.0.0.5522, signed in:**
  - every OPTIONS answer had `X-MSFSSHTTP: 1.3`;
  - the OPTIONS that carried `X-MSGETWEBURL: t` also got `X-MSWEBURL: /`, the server-relative path of the site;
  - Word then POSTed to `<web url>/_vti_bin/cellstorage.svc/CellStorageService`, with no FrontPage or `webs.asmx` request.
- **The spike:**
  - without these headers (runs 1 to 5), Word used plain WebDAV;
  - with them (run 6, `X-MSWEBURL: /n`), Word POSTed to `/n/_vti_bin/cellstorage.svc/CellStorageService` the same request it sent SharePoint.
- **Before sign-in** SharePoint sends neither (see below), so a host must not answer OPTIONS with them unless the request is authenticated, or Word would try FSSHTTP without credentials.

### What Word does when the Cell sub-requests fail

- **Spec:** [MS-FSSHTTP] 2.2.5.6 has `RequestNotSupported` for a sub-request the server does not support.
- **Word,** with `RequestNotSupported` on the Cell sub-requests and success on the rest, sent `ExitCoauthoring` twice. Then:
  - **Run 6 (Windows sign-in, `/n`):** it showed "Sorry, we couldn't open '<url>'", with no GET, HEAD or LOCK.
  - **Run 7 (token in the path, anonymous, `/t/{token}`):** it fell back. It did a GET of the document, then the old discovery at the root of the host (`_vti_inf.html`, the FrontPage RPC, HEAD, `webs.asmx` `WebUrlFromPageUrl`, `workflow.asmx` `GetWorkflowDataForItem`). Then it opened the document read-only, in Protected View.
- **Why they differ:** unknown. The difference between the runs is the sign-in, and the site path.

### The first request differs from the example of MS-FSSHTTP 4.1.1

- **Spec:** [MS-FSSHTTP] 4.1.1 has Coauth, SchemaLock, three Cell, ServerTime and WhoAmI.
- **Word 16.0.20326,** to SharePoint and to the spike alike:
  - **Extra sub-requests:** three EditorsTable `UpdateEditorMetadata` (`OnSuccess` the join; keys `SupportsRename`, `Capabilities`, `Platform`, base64 values), `GetDocMetaInfo` (`OnExecute` WhoAmI), `GetVersions`, and a fourth, 27-byte Cell sub-request with no dependency.
  - **`Timeout`:** `2580`.
  - **`Request` attributes:** `UseResourceID="true"` (plus `ResourceID` when Word knows it from an earlier open), `UserAgent` (a GUID), `UserAgentClient="msword"`, `UserAgentPlatform="win"`, `Build`, `MetaData="1031"`.
  - **HTTP headers:** besides SOAPAction, `Application`, `FileExtension`, `OSName`, `Scenario: OpenFile_Open`, `ScenarioType`, `X-ClientRing` and `X-RequestStats`.
  - **Encoding:** the body is MTOM, but the Cell data is inline base64 (with `BinaryDataSize`), not an XOP part.

### SharePoint's response to it

- **Spec:** [MS-FSSHTTP] 4.1.2 has `ResponseVersion MinorVersion="0"`, `WebUrlIsEncoded="false"`, and `UrlIsEncoded="true"`.
- **SharePoint Server 16.0.0.5522:**
  - **Envelope:** `MinorVersion="3"`; `ResponseCollection WebUrl="https://<host>"` without `WebUrlIsEncoded`.
  - **`Response`:** `Url` unencoded (with spaces) and without `UrlIsEncoded`; it echoes `ResourceID`.
  - **The join:** `TransitionID` equals the `ResourceID`.
  - **Cell sub-responses:** `ServerCorrelationId` (the `SPRequestGuid`). Their data is inline base64 for some and an XOP part (`Content-Transfer-Encoding: binary`) for others, in the same response.
  - **WhoAmI:** `UserLogin` is a claims login (`i:0#.w|<domain>\<user>`); `UserSIPAddress` is empty.

### sharedaccess.asmx is polled from the open

- **Spec:** [MS-SHDACCWS] `IsOnlyClient`.
- **Word:**
  - **Timing:** right after the cellstorage open, it POSTs `/_vti_bin/sharedaccess.asmx` (`SOAPAction: "http://schemas.microsoft.com/sharepoint/soap/IsOnlyClient"`), every 8 seconds for a minute and then every 30 seconds.
  - **Body:** `<IsOnlyClient><id>` is the `ResourceID`, and SharePoint answered `true`.
  - **HTTP headers:** `Scenario: PollingSubscriptionNotConnected_AmIAlone` and `X-RequestStats: Operation=AmIAlone`.


### Required columns block the save before any upload

- **SharePoint Server, a library with required columns:** after an edit, Word's save showed "REQUIRED PROPERTIES: To save to the server, correct the invalid or missing required properties", and sent no cellstorage request.
- **Setting every required column through COM** (`ContentTypeProperties`, each one valid by `Validate()`) did not help. Neither did "Retry Save". Word's property pane said "No SharePoint properties found", so its check fails without the property schema of the library. Word checks this before uploading, with no request to the server.

### A save is a Cell sub-request alone, with BypassLockID

- **Spec:** [MS-FSSHTTP] 2.3.3.1 (CellSubRequestDataOptionalAttributes), and [MS-FSSHTTPB] PutChanges.
- **Word:** a save is one cellstorage request (`Scenario: AutoSaveFile_Save`) with one Cell sub-request. Its attributes are `Coalesce="true"`, `CoauthVersioning="true"`, `GetFileProps="true"`, `BypassLockID` and `SchemaLockID` (both the coauthoring schema lock), and `BinaryDataSize`.
  - **Encoding:** the FSSHTTPB is a binary XOP part, where the Cell queries of the open are inline base64.
  - **`MetaData`:** `7` on the save, `1031` on the open and on `GetDocMetaInfo`.
  - **Afterwards:** Word sends a separate `GetDocMetaInfo` (`Scenario: GetMetadata_EnsureDocMetaInfo`).

### GetDocMetaInfo returns the vti properties of the file

- **Spec:** [MS-FSSHTTP] 2.3.1.27 to 2.3.1.29 define `DocProps` and `FolderProps` as key/value properties without naming them.
- **SharePoint Server:** `DocProps` holds the `vti_*` metadata of the file, mixed with the library's columns. That includes the lock (`vti_sourcecontrollockid`, `…locktype`, `…lockexpires`, `…multiuserchkoutby`), the version (`vti_contenttag`, `vti_contentversionisdirty`, `vti_TimeOfLastCobaltVersionCreation`), and the file (`vti_filesize`, `vti_timelastmodified`, `vti_author`, `vti_modifiedby`, `vti_replid`). The full list is in fsshttp-spike.md.

## FSSHTTPB

The Cell sub-requests of Word below are in the tests as `FssHttpBVectors.WordRequests` (src/OfficeHost.Tests/FssHttp), taken from run 7 of the spike. They are queries, with no document content.

### Word sends a user agent GUID and a client and platform

- **Spec:** [MS-FSSHTTPB] 2.2.2: the User Agent Client and Platform "replaces the User Agent GUID. If this is specified, the User Agent GUID SHOULD NOT be specified."
- **Word:** every Cell request with a user agent has both: the GUID `1984108C-4B93-4EEB-B320-919432D6E593`, then the client `msword` and the platform `win`. The codec reads and writes either or both.

### A lone Query Access request has no user agent

- **Spec:** [MS-FSSHTTPB] 2.2.2 lists the User Agent Start, Version and End as required fields of a request.
- **Word:** its fourth Cell sub-request of the open (27 bytes, `WordRequests[3]`) is a request start, one Query Access sub-request and the request end, with no user agent and no Data Element Package. The codec reads the user agent as optional.

### Query Changes flags of two bytes, with reserved bits set

- **Spec:** [MS-FSSHTTPB] 2.2.2.1.3: the flags A to H take a byte, and I (user content equivalent version OK) is in a second byte with 7 reserved bits that "MUST be set to zero".
- **Word:** sends the two bytes, with the first byte zero and the second `0xAC`, `0x50` and `0x10` in the three queries of the open: reserved bits set, differing from query to query. The codec keeps the flags as sent, and their length, so that it writes back what it read.
- **Excel and PowerPoint** do the same: in the 63 queries of their captures the first byte is always zero, and the second takes values from `0x00` to `0xF8`, its low two bits clear (so never I). None asks for fragments, a file hash or rounded knowledge; see "Query Changes and Put Changes options" for the rest.

### Errata of the examples of MS-FSSHTTPB 4

The bytes of the examples are right where their descriptions are not. The tests use the bytes (`FssHttpBVectors.Spec*`), and each one reads and writes back byte for byte.

| Section | The description says | The bytes are |
|---|---|---|
| 4.1, 4.3 | User Agent Version `0x0008027A` has the type 0x2F | 0x4F, as in the table of 2.2.1.5.2 |
| 4.1, 4.3 | Sub-Request End `0x010B` has the type 0x21 | 0x42, the type of the sub-request start |
| 4.2 | the Waterline `0xFCF808` is 75503, and 4 bytes | 73503, as the cell range before it says, in 3 bytes |
| 4.3 | Put Changes Request `0x002602D2` has the length 9 | 19: two Extended GUIDs of 17 and 1 bytes, then the flags |
| 4.3.5 | the Revision ID is the Extended GUID `14 B9FADE84…`, the Base Revision ID `80 13380CDE…` | those are the root and object of the first Revision Manifest Root Declare (`0x4C50`); the Revision ID is `0C 3AFE2871…` and the Base Revision ID is null (`00`) |
| 4.4 | the signature is `0x9B069439F329CF9C`, of a request | `9D CF 29 F3 39 94 06 9B`: `0x9B069439F329CF9D`, of a response (2.2.3) |

Two more places where the spec contradicts itself:

- **Query Changes Request (0x51):** the table of 32-bit starts (2.2.1.5.2) has it as not compound, and the example of 4.1 has no end for it, but the table of 16-bit ends (2.2.1.5.4) lists it. The codec treats it as single.
- **Put Changes Response (0x87):** 2.2.3.1.3 lists its header as a required field, with only the applied Storage Index and the data elements added optional; the example of 4.4 has a Put Changes sub-response without it, the knowledge right after the sub-response start. The codec reads the header as optional, and writes it only when the model has it.


## MS-FSSHTTPD

The evidence here is the three Put Changes requests of Word's saves to SharePoint Server (17 to 20 KB each, the binary part of `Scenario: AutoSaveFile_Save`, see the saves above). They hold the document's content, so they are not in the repository: they were checked locally with the codec, which reads each and writes it back byte for byte.

### Word's signatures: exclusive or, with no CRC of the compressed data

- **Spec:** [MS-FSSHTTPD] 2.4.1: the signature of the data of a zip item is its CRC and sizes. The note <6> says Office adds the CRC of the compressed data after the CRC from the Office 2010 March 2015 CU. For protocol version 2.2 and later, the signature of an item of at most 4096 bytes is the exclusive or of those of its header and data.
- **Word 16.0.20326:** the data signatures are 20 bytes, the CRC and the two sizes, with no CRC of the compressed data. A small item's signature is the SHA-1 of its header exclusive or that, 20 bytes. A header chunk is the SHA-1 of the header, and the central directory the SHA-1 of its bytes. All 27 leaves of the three saves whose data node is in the same request match `ChunkingOptions.Default` (exclusive or, no CRC of the compressed data), which is what this host writes.

### How Word lays out the tree

- **Spec:** the example of [MS-FSSHTTPD] 3.1 has one object per Object Group, no Object Group metadata, and leaves with no Data Hash.
- **Word:**
  - each leaf is in one Object Group with its Data Node Object, and each Object Group has Object Group metadata declarations ([MS-FSSHTTPB] 2.2.1.12.6.3), one per object;
  - every leaf has a Data Hash of 20 bytes, whose algorithm the note <3> leaves incomplete ("without the introduction of the length"), so this host keeps those it reads and writes none;
  - a revision has no base revision: its Revision Manifest lists the Object Groups it keeps from earlier saves, by their identifiers, with its new ones. `FileCell.Write` does the same.
- This host writes the layout of the example, which its tests reproduce data element for data element. Word reads layouts other than its own, as it reads SharePoint's: it opens the revisions the service writes from a file (a new file, or one changed outside Word).

### A second root: an opaque stream

- **Spec:** [MS-FSSHTTPD] 2.3: "If other streams are present, additional roots MUST be specified by the protocol client and are opaque to this protocol."
- **Word:** every revision has a second root, `{EDF48BAC-3745-4416-A4C9-D5B71BB870F6},2`, a single leaf of 199 bytes with an empty signature over a Data Node Object. `FileCell.Write` keeps the other roots of the previous revision with their Object Groups.

### Serial Numbers and the expected Storage Index

- **Word:** every data element of a Put Changes has the null Serial Number; the server assigns them. The mappings of the Storage Index do have Serial Numbers, of a GUID of Word's. The Data Element Package also holds the Storage Index of `ExpectedStorageIndex`, with its Storage Manifest and Cell Manifest. The service assigns the Serial Numbers, and compares the expected Storage Index with its own (fsshttp-concurrency.md, "Coherency").

### The order of the 32-bit Extended GUID

- **Spec:** [MS-FSSHTTPB] 2.2.1.7.5 has the type 128, then the value (4 bytes), then the GUID. The description of the example of [MS-FSSHTTPD] 3.1 decodes `80 EC BC 97 4D DC 28 … 6F 17 01 00 00 11` the other way, as the GUID `{4D97BCEC-28DC-41C5-9274-26CB57966F17}` and the value `0x11000001`, and the bytes of both examples look made that way (the values `0x11000001` and `0x12000002` to `0x12000007` under one GUID).
- **Resolution:** the codec follows 2.2.1.7.5, as Microsoft's own protocol test suite does (OfficeDev/Interop-TestSuites, `ExGuid.cs`, MIT; read, nothing copied). Word's saves use no 32-bit form (their values are below 0x20000), so no capture settles it. This host only writes the 32-bit form after 131,071 identifiers in a file.

### Gaps in the spec

- **Leaves of 1 to 3 megabytes:** 2.4.1 creates a Data Node Object only for leaves of at most 1 megabyte, but chunks are subchunked only above 3 megabytes, so a chunk in between has no defined data. This host gives every leaf a Data Node Object.
- **RDC analysis (2.4.2):** not done. A file that is not a zip is chunked by megabyte (2.4.3). The tree of any chunking is read the same, so files Word chunks by RDC are read.
- **Unique signatures:** the large final chunk and the subchunks need signatures that are only unique (Office makes them random, notes <7> and <9>). This host uses the start of the SHA-1 of the position and bytes of the chunk, so that an unchanged chunk keeps its signature, and its objects.


## Word against OfficeHost

Word 16.0.20326, on the `desktop` launch profile of the sample (`https://localhost:5001`), with every request logged. It opened the document, enabled editing, saved twice and closed, and the file had both edits.

The expected Storage Index of its saves, and the refresh of its lock, are in fsshttp-concurrency.md.

### Word opens from its cache only when the answer has the Storage Index

- **Spec:** [MS-FSSHTTPB] 3.1.4.2 has the server send "a full or partial set of data elements"; nothing says which
  ones a client that knows them all gets.
- **Word** (2026-11-05, documents of 13 MB and of 320 KB it had saved): its join's content query had the knowledge of
  its cache, and OfficeHost answered with nothing, the client knowing every data element. Word then queried the content
  and the metadata again, with empty knowledge, and downloaded the whole file: every open was a full download (step 6
  of docs/fsshttp-checklist.md had not passed), and one over Word's maximum failed (see the next section).
- **SharePoint Server 2016** sends a Storage Index in every Query Changes answer of the recordings (30 of 30, the
  three partitions, from the open to the polls), with knowledge one Serial Number ahead of the client's.
- **The service** sends the Storage Index in every complete answer, known to the client or not. Word then opened
  from its cache, with no second query. The replay did not show the difference: it compares the content of the cells,
  not the data elements of each answer.

### A download over Word's maximum comes in parts

- **Spec:** [MS-FSSHTTPB] 2.2.2.1.3: the Maximum Data Elements is "the limit of data elements at which the server
  starts breaking up the results into partial results", and 2.2.3.1.2 has the partial result flag. Neither says what
  a partial result holds, nor what the client does next.
- **Word** asks for 3,670,016 bytes at most, and queries again after a partial answer, with the knowledge of that
  answer, until one is not partial.
- **SharePoint Server 2016** (2026-11-06, a document of 7.5 MB, Office's cache cleared): three answers, of 3.1, 3.1 and
  1.2 MB, within a second. The two partial ones have the flag, a null Storage Index Extended GUID, Object Groups only
  (of 1 MB at most), and as cell knowledge the Serial Numbers of the Object Groups sent so far, one by one (and
  `0..0`); their file properties have no `Etag`. The last has the other Object Groups, the manifests and the Storage
  Index, its Extended GUID, the whole knowledge, and the `Etag`. A document of 38 MB, the same way: 12 partial
  answers of 3.1 to 3.6 MB and a last of 0.6 MB, within 14 seconds.
- **The service** named its Storage Index in a partial answer, and sent the `Etag`: Word sent no other query, exited
  coauthoring, and the open failed, or Word fell back to a GET of the file and opened it read-only (with the Storage
  Index sent too, in an experiment, the same). It now answers as SharePoint: no Storage Index named in a partial
  answer, no `Etag` (`FssHttpServiceTests.Query_Partial`). Word then opened a document of 13 MB it had not cached in
  four answers of 3.3, 3.3, 3.3 and 3.2 MB, for editing, with its pictures, and one of 43 MB in 14 answers within 5
  seconds.

### Protected View, then a second open

- **Word:** on a host in the Internet zone, the open (join, queries) ends with `ExitCoauthoring` (`Scenario: ReleaseLock_LockRelease`) and the document in Protected View. Enable Editing sends the open again, with the join, and the editing goes on from there, polling `IsOnlyClient`.

### What Word does without

- `GetVersions` answered `RequestNotSupported`, `POST {site}/_api/web/GetFileByUrl(@v)/ListItemAllFields/GetSharingInformation` and `POST /_vti_bin/webs.asmx` at the root answered 404: Word opened, edited and saved all the same.
- The editors table partition and the metadata partition answered with nothing in them (a Storage Index with a null Storage Manifest mapping), and the open went on.

### The second root is a stream like the editors table

The 199-byte stream of Word's second root (`{EDF48BAC-…},2`, above) starts with `1A 5A 3A 30 00 00 00 00`, the Editors Table Zip Stream Header of [MS-FSSHTTPB] 2.2.3.1.2.1.1: a DEFLATE-compressed XML of Word's.

## Excel

Excel 16.0.20326 (`UserAgentClient="msexcl"`), opening a workbook with the edit button of the application on SharePoint Server, changing a cell, saving and closing (captured with `src/fsshttp-capture.ps1 -HostName`); then the same against the `desktop` launch profile of the sample. Its exclusive lock, and a second user locked out of the workbook, are in fsshttp-concurrency.md ("Excel: the exclusive lock").

### GetDocMetaInfo has vti_canmaybeedit

- **SharePoint:** among the properties of `GetDocMetaInfo`, `vti_canmaybeedit=true`, and, once locked, `vti_sourcecontrollockid` (the exclusive lock id, without braces, in lower case) with `vti_sourcecontrollocktype=0` and the user in `vti_sourcecontrolcheckedoutby`.
- **The service** sends `vti_canmaybeedit` (false without write access) and the lock id and type. Excel locked and saved with it; whether Excel needs it was not isolated, as the tests without it were through COM.

### GetVersions

- **SharePoint** answers it with the versions of the file (`<results><list/><versioning enabled="1"/>…`). The service answers `RequestNotSupported`, and Excel opened, edited and saved all the same.


## SharePoint Server SE against OfficeHost, request by request

`src/sharepoint-lab/Compare-CellStorage.ps1` sends the same sub-requests to the lab (SharePoint Server Subscription Edition 16.0.14326, as `LAB\alice`) and to the sample, and shows where the answers differ. The scenarios come from the server rules of Microsoft's Interop-TestSuites (`MS-FSSHTTP-FSSHTTPB_RequirementSpecification.xml`). OfficeHost answers as SharePoint SE does, including these, which the specs do not settle:

| Request | SharePoint SE and OfficeHost |
|---|---|
| ExitCoauthoring again, after the session ended | `Success` |
| ExitCoauthoring under another schema lock id | `Success`, and the client stays in the session with its lock |
| GetCoauthoringStatus without a session | `CoauthStatus="None"` |
| MarkTransitionComplete without a session | `Success` |
| ExclusiveLock ReleaseLock without a lock | `Success` |
| `OnExecute` on a sub-request kept from running by a failed `OnSuccess` or `OnFail` | `DependentRequestNotExecuted` |

- **`OnExecute`** runs after a dependency that ran and failed, and after one kept from running by `OnNotSupported`: Word's open has its Cell queries `OnExecute` a SchemaLock kept so.
- **Requests Office does not send** (none of the captures has an `Etag` on a Cell sub-request, or `Properties`), answered alike too:

  | Request | SharePoint SE and OfficeHost |
  |---|---|
  | A join, schema lock or exclusive lock with a `Timeout` of 120000 or more (119999 is accepted, 30 too) | `LockRequestFail` |
  | A Cell sub-request with the `Etag` of another version | `CellRequestFail` |
  | Properties `PropertyGet` | `Success`, `<PropertyValues/>` |
  | Properties `PropertyEnumerate` | `RequestNotSupported` |
  | `AmIAlone` | `True` or `False`, capitalized |

- **Also alike:** `AmIAlone` with the session's `TransitionID` or another one (it answers by the number of clients); twelve coauthors (no limit reached); a Cell query with `ExpectNoFileExists` on an existing file (`Success`).
- **Office quirks reported by cellbridge** (a similar MIT-licensed host): Word rejecting an `ErrorStringSupplementalInfo` in UTF-8 (the codec writes string items in UTF-16, as [MS-FSSHTTPB] 2.2.1.4); Word turning read-only when the content query answers an empty Storage Index (the content partition always holds the file); Word resuming an upload without a Query Changes, with a stale expected Storage Index (a coherency failure, as on SharePoint, after which the client queries).
- **Where SharePoint fails** (a 302 to its error page, "Object reference not set to an instance of an object"), OfficeHost keeps its answers: a Coauth or ExclusiveLock without its request type (`InvalidArgument`), `AmIAlone` without data, a dependency on an unknown token, a `RequestVersion` of 1.
- **Different:** SharePoint supports `GetVersions` (OfficeHost: `RequestNotSupported`, see Versions below), and adds an `ErrorMessage` to the `InvalidArgument` of a join without a schema lock id.

### FileOperation Rename

- **Spec:** [MS-FSSHTTP] 2.3.1.34: `FileOperation="Rename"`, `NewFileName`, and an optional `ExclusiveLockID` that lets the client holding the exclusive lock rename.
- **SharePoint Server SE and OfficeHost** (`FileOperationTests`):

  | Request | Answer |
  |---|---|
  | Rename, no lock | `Success` |
  | Rename to another extension (`.docx` to `.txt`) | `Success` |
  | Rename with nothing in `NewFileName` | `Success`, nothing renamed |
  | Rename in a coauthoring session (even by its client) | `FileAlreadyLockedOnServer`, with no `ErrorMessage` |
  | Rename under an exclusive lock, with its id / without | `Success` / `FileAlreadyLockedOnServer` |
  | Rename to an illegal name | `InvalidArgument` |
  | Rename to the name of another file | `FileAlreadyExists` (SharePoint's `ErrorMessage` is the text of its exception) |
  | No `FileOperation` | SharePoint fails; OfficeHost `InvalidArgument` |

- **OfficeHost** renames through `IOfficeStorage.Rename`, with the `CanRename` of the access, as WOPI's RenameFile. The site of the file keeps working under its new name, the token in its path standing for the file. A request to the old name fails with `FileNotExistsOrCannotBeCreated` (a GET or HEAD with a 404), as on SharePoint: the name in the URL picks the file since Save As (see "New files" below).

### Versions

`GetVersions` (Word, Excel and PowerPoint send it in every open) and `Versioning` `GetVersionList`, against the lab (versioning on, two versions of `Lab.docx`):

- **`GetVersions`** answers [MS-VERSS] 2.2.4.1 `results`: `<list id>`, `<versioning enabled="1"/>`, `<settings url>`, and a `<result>` per version: `version` (`@2.0` for the current one, `1.0` for an older one), `url` (the current one at the file's URL, older ones at `/_vti_history/{512 × major}/{path}`), `created` (a local display time), `createdRaw` (UTC, ISO 8601), `createdBy` (the login), `createdByName`, `size`, `comments`.
- **`Versioning` `GetVersionList`** answers a `UserTable` (`User` with `UserId`, `UserLogin`, `UserName`, `UserEmailAddress`) and `Versions`: `Version` with `Number`, `IsCurrent`, `LastModifiedTime` (FILETIME) and `UserId`, the current one with `Events` (`Event` `Id`, `Type="2"`, `CreateTime`, `UserId`), one per save.
- **What OfficeHost would need from the application:** the versions of a file (an id or label, when, by whom, the size, a comment, and which is current), the content of an older one (for its URL, under the desktop site), and, for `RestoreVersion`, restoring one. An optional store, as `IOfficeUserInfoStore`: the application decides what makes a version. Office opened, edited and saved without any of it.

### Query Changes and Put Changes options

What Office sets, from the 113 Cell sub-requests the sample logged (Word, Excel and PowerPoint, opens and saves; 71 Query Changes, 4 Put Changes) and Word's saves before them:

- **Query Changes:** the flags above, the arguments `3` (include the Storage Manifest and the cell changes) with a null Cell ID, at most 3,670,016 bytes of data elements (the service answers partial results beyond it), the client's knowledge, no versioning and no filters.
- **Put Changes:** flags `0x49`: A, imply null expected if no mapping; D, favor coherency failure over not found; G, return complete knowledge (Excel once sent `0x09`). Then three zero bytes: an empty content version coherency check, no author logins, the reserved byte. Additional flags `0x8004`: C, check for id reuse, and bit 15, which the spec reserves. No lock id, the client's knowledge, no diagnostic. A save Word splits (see "SharePoint Server 2016") has `0x4B` for its parts (B, partial, besides) and `0x0D` for the last (A, C partial last, D), with G on some only.

SharePoint Server SE and OfficeHost answer each case of `PutChangesTests` (a save as Office's, with one thing changed) alike. The cases of the expected Storage Index, and how coherency is checked, are in fsshttp-concurrency.md ("Coherency is key by key").

| Put Changes | SharePoint SE and OfficeHost |
|---|---|
| A data element the new Storage Index reaches is in neither the package nor the storage (with D or without) | `CellRequestFail`, Cell error 16 |
| A data element under the id of another the storage has, C / no C | Cell error 112, Extended GUID collision / applied |
| A failed Put Changes, in the SubResponse | `CellRequestFail`, the Cell error in the FSSHTTPB |

- **D made no difference:** SharePoint answered 16 for a missing data element whether it was set or not, and for an expected Storage Index missing from the package.
- **The Put Changes response** of SharePoint has 29 bytes the spec does not define, inside the response object after the data elements added: a GUID that was the same on every save of the run, `7A 05 24 00 03 05 44 00`, then a number that grew from save to save. Office saves to OfficeHost without them; the codec keeps them as `Extra`. SharePoint also returns the data elements added without B being set.

### Word's session replayed (docs/replay.md)

Word 16 (through COM) opened `Lab.docx` of the lab as `LAB\alice`, added a paragraph, saved and closed, with the replay recorder in SharePoint. Its replay to OfficeHost found these; OfficeHost answers them as SharePoint does, but for the timeout:

- **Query Access** has a third access response after the write access response, of stream object type 0xA4, which the spec does not define (the type tables of [MS-FSSHTTPB] 2.2.1.5 have none): E_ACCESSDENIED chained to E_ACCESSDENIED, for alice with Contribute rights. Read and write were allowed. The codec keeps the third as `ExtraAccess` (`FssHttpBVectorTests.SharePoint_QueryAccessResponse`). OfficeHost sends none, and Word opens and saves without it.
- **The editors table has the time of the join:** SharePoint adds `JoinCoauthTime` to the metadata of the editor at its join, a FILETIME in base64 (05:42:18.077 for a join at 05:42:18), before the client's own metadata (Word's open sends `SupportsRename`, `Capabilities` and `Platform` after its query of the editors table). OfficeHost adds it too, outside the client's quota of four pairs.
- **`HaveOnlyDemotionChanges`** is in the file properties (`GetFileProps`) of a query that sends changes: not of a query that has none (the knowledge only, about 200 bytes), nor of a put. OfficeHost answers so. Seen on 7 content queries: with it, 589 to 16,432 bytes; without, 149 to 222.
- **The editors table's timeout** is the join plus the client's `Timeout` on SharePoint (2,580 seconds), where OfficeHost keeps a client an hour at least and rounds the time up to ten minutes. Not changed: the spec lets the server pick its own from 60 to 3,600 seconds.
- **IsOnlyClient** came without a user: `sharedaccess.asmx` is served without sign-in on SharePoint.

A second recording added Word with a second client in its session, a Save As, Excel and PowerPoint (docs/replay.md):

- **RefreshCoauthoring** is answered `LockType` and `CoauthStatus`, without the `TransitionID` of a join. OfficeHost sends the `TransitionID` with a join only.
- **The Cell of a Save As** (its `ExclusiveLockID`, `Timeout` and a whole file) is answered `LockType="ExclusiveLock"` with its file properties: it holds the lock it took. OfficeHost answers so.
- **`GetFileProps` is for the content:** a Cell of the editors table with `GetFileProps="true"` (Word's after a Save As) has no `Etag` or times. OfficeHost gives the properties to the Cells of the content only.
- **A second client joining** makes Word's next IsOnlyClient `false`, then Word sends `MarkTransitionComplete` with its three Cell queries, then a poll of three Cell queries and `GetCoauthoringStatus` every 30 seconds while coauthoring, and IsOnlyClient again once alone. OfficeHost answered all of these as SharePoint did.
- **Excel through COM** opened `Lab.xlsx` read-only (no lock, nothing saved), with SharePoint's `vti_canmaybeedit` true and no lock on the file: Excel's choice, as "Excel takes an exclusive lock at the first edit" (fsshttp-concurrency.md) says of its object model.
- **The same as SharePoint:** the join, the schema lock kept from running, the editor metadata, WhoAmI, ServerTime, the IsOnlyClient polls, the exit and the release of the lock. Different as known: GetVersions (see Versions), the `vti_` properties of GetDocMetaInfo (OfficeHost sends a subset), and the Cells of documents Office had in its cache before the recording, which the replay cannot translate.

A third recording, of documents Office had never opened, with the cells translated (docs/replay.md, "The cells"):

- **Word's and PowerPoint's puts** (two saves, a save after opening again, PowerPoint's save) were applied by OfficeHost as by SharePoint, to the same file.
- **Opening a document again,** after its own saves, Word was sent changes by SharePoint (`HaveOnlyDemotionChanges` `False`) that left the file as it was: something of SharePoint's own in the cell, as a new revision. OfficeHost had nothing new to send, and Word opened, edited and saved as well.
- **PowerPoint's first request,** `GetMetadata_EnsureDocMetaInfo` on a keep-alive connection, came with no body in two recordings though it declared a `Content-Length`: PowerPoint dropped the connection (IIS: `sc-win32-status` 64), SharePoint answered its error page to no one, and PowerPoint sent the request again 30 ms later, which succeeded. It came from the replay recorder reading the request body at authentication: without the recorder (the IIS log of 2026-10-21: 12 requests of PowerPoint, all 200), and with the recorder reading it after SharePoint's handler, as it does (9 requests of PowerPoint, from a new process, all 200), every request had its body (docs/replay.md). Not a behaviour of PowerPoint's.
- **SharePoint SE serves `cellstorage.svc`** with `/_vti_bin/cobalt.ashx/cellstorageservice`, an ASP.NET handler, not WCF: its IIS log has the 401 of the handshake under the first, and the request signed in under the second.

### New files

Word 16.0.20326 against the lab and the sample, driven through COM (`SaveAs2`), captured with `fsshttp-capture.ps1`.

- **An empty file** (0 bytes, as an application makes for a new document): SharePoint serves it as any file, and Word opens a blank document; against the sample too, which syncs the empty stream as a cell (`NewFileTests.EmptyFile`). Its first save is an ordinary Put Changes. Office for the web does the same through WOPI: the `editnew` action of the discovery, on an empty file, whose first PutFile needs no lock ([MS-WOPI] 3.3.5.3.2, `ApiTests.Launch_EditNew`).
- **Save As to the same library**, against the lab:
  1. a FrontPage RPC (`application/x-vermeer-rpc`, the web of the new URL) and `lists.asmx` `GetListContentTypesAndProperties`, through WinINet; against the sample, whose OPTIONS announces no FrontPage, neither;
  2. `HEAD` of the new URL: 404;
  3. one cell storage Request to the new URL (`X-RequestStats: Operation=SaveAs`): a Cell sub-request whose Put Changes has the whole file, a new Storage Index and no expected one, flags `0x48` (no imply null), additional flags `0x0004`, with `Coalesce`, `CoauthVersioning`, `GetFileProps`, and `BypassLockID` and `ExclusiveLockID` of a new id with `Timeout="600"`; then, on its success, ExclusiveLock `ConvertToSchemaJoinCoauth` (and `ConvertToSchema` on not supported), the editors table metadata, a Cell query of the editors table, `WhoAmI`, `GetDocMetaInfo`, `GetVersions` and `ServerTime`. SharePoint answers all `Success`: the Cell created the file and took the exclusive lock, which the ExclusiveLock turned into a session;
  4. Word goes on with the new file: IsOnlyClient with its ResourceID as the `id`, its saves, and the lock release of the old one.
- **OfficeHost** (`NewFileTests`): the name in a document URL of the site (`{site}/{name}`, the Request `Url`, the path of a GET or HEAD) picks the file: the file of the token under its current name, or a file the site made by Save As. With `CanWriteRelative` (and `CanWrite`), as the Save As of WOPI, a Request with a Put Changes to a free name creates the file next to the file of the token (`IOfficeStorage.CreateSibling`), or takes over an unlocked file of that name, which it writes but does not read: its cell is written again with new ids, a query of its content in the Request fails with `StorageWriteOnly`, and the site reaches it only once the Put succeeded (`NewFileTests.SaveAs_OtherFile_NotRead`); the Cell takes the exclusive lock its `ExclusiveLockID` and `Timeout` name. A created file whose Put fails is deleted again. The files a site made are kept with the state of the file of its token (`SaveAsFiles`, a partition of the `IFssHttpStore`), by a hash of the token until it expires, so the site reaches those and no other file of the folder; IsOnlyClient finds them by ResourceID. Only the Open XML formats are created; another name answers `FileNotExistsOrCannotBeCreated`, as does any name without `CanWriteRelative`. A HEAD of a file the site could take over answers 200 without the content, so that Word asks before replacing it.
- **Against the sample**, with tokens of 12 characters, Word saved as `SavedAs.docx`, then edited and coauthored it through the site. Word refuses a document URL of more than 255 characters ("String is longer than 255 characters"), which a JSON Web Token of about 270 characters in the path exceeds: so the built-in tokens are compact (`SignedToken`, about 70 characters), and Word saves as through them too.
- **The cell partitions** of OfficeHost's own state (the coauthoring session, the Save As files) are refused to Cell sub-requests with `InvalidPartition`: a client could otherwise read or forge them.

### SharePoint wants a line break before the first MIME boundary

- **SharePoint:** an MTOM request whose body starts with the boundary, without the CRLF Word puts before it, fails with the same 302 to its error page, before the cell storage service runs. Word's request with its sub-requests replaced works; the headers (`User-Agent`, `X-Office-*`, `Scenario`) and the `Request` attributes make no difference.
- **OfficeHost** reads both.


## SharePoint Server 2016

Word against a SharePoint Server 2016 farm (16.0.5522), recorded from inside it by the module of the shadow
(docs/shadow.md) and replayed to OfficeHost (docs/replay.md): a new document opened from a library, edited, saved,
left open 24 minutes, closed. 10 exchanges, 47 sub-responses: OfficeHost had the same file as SharePoint after each of
the 7 Cells of the content, the save too, and answered 30 sub-responses as SharePoint did. The other 17 are where 2016
answers otherwise than SE, which OfficeHost follows, now known differences (`ShadowRule.Defaults`):

- **`cellstorage.svc` is WCF,** not SE's `cobalt.ashx`: WCF reads the request body unbuffered, and its response is
  flushed write by write (what the module of the shadow had to be made for, docs/shadow.md).
- **`HaveOnlyDemotionChanges="False"`** is on every Cell of the content: the queries, with changes or none, the 27-byte
  Query Access of the open, and the put. Never on the other partitions. SE sends it with a query that sends changes.
- **The editors table** has no `JoinCoauthTime`, and the `EmailAddress` of the editor. Once Word updates its metadata
  (`UpdateEditorMetadata`), the editor has no `HasEditorPermission`. Its `Timeout` has the ticks of the join.
- **MarkTransitionComplete** of a client that has left the session (after its ExitCoauthoring) is
  `InvalidCoauthSession`, where SE answers `Success`. Word opened the document again just after, with a join.
- **No editors table** in the answer to the query of the editors partition when no one is editing, where OfficeHost
  sends an empty one.
- **A save of more than about 3 MB of changes is split** into partial Put Changes ([MS-FSSHTTPB] 2.2.2.1.4, Partial
  and Partial Last), seen with pictures added to a document: 3.1 MB, 3.1 MB, then 1.2 MB, within 5 seconds. Each part
  has Object Groups only, and a Storage Index of its own that maps nothing, with no expected one (flags `0x4B`). The
  last (`0x0D`) has the other Object Groups, the manifests, and the Storage Index of the save with the expected one.
  SharePoint answers each part as a Put Changes that applied nothing: Data Elements Added empty, knowledge without
  cell knowledge, and the file properties without `Etag`; the last as a save. The spec says nothing of how a server
  ties the parts together. OfficeHost failed each part (`InvalidObject`: it wrote the file of each put), and now keeps
  their data elements by file, user and partition (`PartialPuts`), each part for an hour (those of a save that never
  ended leave the store with the file's next Put Changes), answers them as SharePoint does, and applies them with the
  last (`FssHttpServiceTests.Save_InParts`). Nothing in a part names its save, so a last part takes from those waiting
  the data elements its save holds, and leaves the others to the last part of their own save: two saves of one user
  at once, as of Word on two machines, do not lose each other's parts. The parts waiting for a file are bounded
  together by `MaxFileSize`, as a save in one request is: the part beyond it is answered 413
  (`Save_InParts_OverMaxFileSize`). Forgetting the parts comes after the save is applied, and its failure does not
  fail the save. Against the sample (2026-11-05), Word saved three
  pictures of 4.3 MB in four parts of 3.1, 3.1, 3.1 and 3.6 MB, flags `0x4B`, `0x0B`, `0x0B` and last `0x4D` (G,
  return complete knowledge, only on some): each answered `Success`, the `Etag` only after the last, and the file had
  the pictures byte for byte. Larger (2026-11-06): SharePoint 2016 took a save of 31 MB of pictures in 10 parts within
  19 seconds; against the sample, Word saved 30 MB in 10 parts within 5 seconds, from 13 to 43 MB.
- **Word's coauthoring** went otherwise than against SE: in a first recording, seven pairs of
  `GetChangesFile_MarkCoauthTransitionCompleteAndDownload` and `PollingSubscriptionNotConnected_TransitionToCoauthInHostMode`
  within 1.5 seconds of the open; in the second, one pair, then ExitCoauthoring and ReleaseLock, and the open again.


## PowerPoint

PowerPoint 16.0.20326 (`UserAgentClient="msppt"`), opening, editing (text appended to a slide through COM), saving and closing a presentation, against SharePoint Server (the lab of docs/sharepoint-lab.md) and against the `desktop` launch profile of the sample. It worked against the sample as it was, with no change to the service.

### PowerPoint coauthors as Word does, with a GetDocMetaInfo first as Excel does

- **The open:** `WhoAmI` and `GetDocMetaInfo` (`Scenario: GetMetadata_EnsureDocMetaInfo`), as Excel; then `OpenFile_Open` as Word's: Coauth `JoinCoauthoring` (Word's `SchemaLockID`, `Timeout="2580"`), SchemaLock `GetLock` `OnNotSupported` it, a Cell query of the editors table partition and one of the content, three `UpdateEditorMetadata` (`Platform`, `Capabilities`, `SupportsRename`), `GetVersions`, `ServerTime` and a Query Access. No query of the metadata partition, which Word makes.
- **The join** also carries `AllowFallbackToExclusive="true"` with an `ExclusiveLockID`: the exclusive lock PowerPoint would take if coauthoring were not supported. The service coauthors, and ignores it.
- **Then** the `IsOnlyClient` polls (a burst every 8 seconds, as Word), `AutoSaveFile_Save` (one Cell sub-request), and `ExitCoauthoring` with SchemaLock `ReleaseLock` on close. Against the sample, the Protected View open and exit before Enable Editing, as Word.
- An edit through PowerPoint's object model saved through the service, unlike Excel's.


## Authentication

### MS-OFBA is blocked by default, for SharePoint too

- **Spec:** [MS-OFBA] 2.2.1 has the client identify itself with `X-FORMS_BASED_AUTH_ACCEPTED: t` (or a listed user agent). On a 403 with `X-FORMS_BASED_AUTH_REQUIRED` (2.2.2), the client MUST navigate to the login page.
- **Word:** it sent no `X-FORMS_BASED_AUTH_ACCEPTED`. On the 403 of run 3 it did not open the login page. It showed "Office has blocked this content because it uses a sign-in method that may be insecure. You can choose to not block such content in the future by changing your settings in the Office Trust Center."
- **SharePoint Server 16.0.0.5522 (forms-based sign-in), captured with `src/fsshttp-capture.ps1`:** Word sent the same OPTIONS. SharePoint answered 403 with `X-Forms_Based_Auth_Required`, `X-Forms_Based_Auth_Return_Url`, `X-MSDAVEXT_Error: 917656; Access denied. Before opening files in this location, you must first browse to the web site and select the option to login automatically.` and `X-MS-InvokeApp: 1; RequireReadOnly`. Word showed the same dialog and sent nothing more.
- **The setting:** File > Options > Trust Center > Trust Center Settings > Form-based Sign-in. Its Group Policy is "Control how Office handles form-based sign-in prompts". By default it blocks every prompt; it can be set to allow listed hosts.

### Word's forms sign-in dialog is the IE engine

- **Spec:** [MS-OFBA] 3.2 leaves the login dialog to the client.
- **Word:** with forms-based sign-in allowed for the host ("Ask me what to do for each host", then Yes), Word opened the login page in a `TridentContainer` window. That is the IE engine (MSHTML), and its requests go through WinINet inside the Word process. They carry the IE `Accept` header (`image/gif, image/jpeg, …, application/x-ms-xbap`).
- **SharePoint Server's login page** lists "Windows Authentication" and "ADFS". Both failed on a PC that is not on the SharePoint domain:
  - **Windows:** `POST /_login/default.aspx` then 302 to `/_windows/default.aspx`. The 401 `WWW-Authenticate: NTLM` led to an NTLM type 2 challenge, and the dialog stopped there.
  - **ADFS:** 302 to `/_trust/default.aspx?trust=ADFS`, 302 to `https://<adfs>/adfs/ls/?wa=wsignin1.0&wtrealm=urn:…`. ADFS answered with a 302 to `/adfs/ls/wia`, its Windows-integrated endpoint, because it treats IE as a browser that can sign in with Windows. Chrome gets the ADFS password form instead.

  In both cases the dialog showed "Can't reach this page" (titled `javascript:` the first time) and never asked for a password.
- **What it means for a host:** forms-based sign-in only works in Word if the login page works in the IE engine. Its identity provider must not send IE to Windows sign-in on PCs outside the domain.

### Word calls the SharePoint REST v2 API before sign-in

- **Spec:** none: `/_api/v2.0` is the SharePoint REST API (OneDrive API shape), not an Open Specification.
- **Word:** after the 403 of the OPTIONS, it sent `GET /_api/v2.0/shares/u!<base64url of the document URL>/driveitem/?$select=viewpoint` through WebIO (WinHTTP). SharePoint Server answered `401 WWW-Authenticate: NTLM`.

### SharePoint sends no capability headers before sign-in

- **SharePoint Server 16.0.0.5522:** the 403 answer to an unauthenticated OPTIONS has no `DAV`, `MS-Author-Via`, `Allow` or web URL; only `MicrosoftSharePointTeamServices` and the sign-in headers. So [MS-WDVME] 3.2.5.1 (which advertises capabilities on OPTIONS) applies only after sign-in, and so does any answer to `X-MSGETWEBURL`.

### Basic over HTTPS is blocked by default

- **Spec:** none: Basic is [RFC7617], and the `Accept-Auth` Word sends lists `Basic`.
- **Word:** on the 401 `WWW-Authenticate: Basic` of run 4 it showed no credential prompt. It showed "Microsoft Office has blocked access to "https://localhost:5001/b/" because the source uses a sign-in method that may be unsecure."

### Windows auth is silent

- **Word:** on the 401 `WWW-Authenticate: Negotiate` of run 5, it authenticated with no prompt and went on as an anonymous client would.


## WebDAV

### LOCK and UNLOCK around a read-only open

- **Word:** in Protected View (all runs on `https://localhost`, which is in the Internet zone), it took an exclusive LOCK (`Timeout: Second-600`, owner `AzureAD\<user>`), did a GET and a PROPFIND, and then an UNLOCK within a second. After "Enable Editing" it repeated the sequence and stayed in Protected View, with a "We've refreshed this document with the latest changes from others" bar.
- **Open:** whether that is the fake LOCK answer of the spike or the zone. Retest with the host in the Local intranet zone.

[MS-FSSHTTP]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/05fa7efd-48ed-48d5-8d85-77995e17cc81
[MS-FSSHTTPB]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/f59fc37d-2232-4b14-baac-25f98e9e7b5a
[MS-FSSHTTPD]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpd/
[MS-WDVME]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wdvme/8cafdf55-ee5c-443e-bdb7-2cb2ab1fb2c3
[MS-WEBDAVE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-webdave/9fb175f3-f063-4389-a959-4457692e4b49
[MS-OFBA]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-ofba/30c7bbe9-b284-421f-b866-4e7ed4866027
[MS-FPSE]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fpse/1bbae50a-7de9-45ea-830d-4e9caf8dfccd
[RFC7617]: https://www.rfc-editor.org/rfc/rfc7617

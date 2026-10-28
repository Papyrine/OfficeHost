# Replay

The replay compares OfficeHost with SharePoint on real traffic. A recording of what desktop Office sent SharePoint, and
what SharePoint answered, is replayed to an OfficeHost of its own, request by request. Then each answer is compared
with SharePoint's. [Compare-CellStorage.ps1](sharepoint-lab.md) sends sub-requests written by hand; the replay sends
the ones Office sent.

The engine is in [src/OfficeHost.Shadow/Replay](../src/OfficeHost.Shadow/Replay) (`ShadowReplay`, `ShadowComparer`,
`ReplayCells`), shared with the live shadow of [shadow.md](shadow.md). `Replayer`
([src/OfficeHost.Tests/Replay](../src/OfficeHost.Tests/Replay)) replays a recording with it to an OfficeHost on a test
server, and `ReplayTests` runs it:

```bash
$env:OFFICEHOST_REPLAY = 'D:/SharePointLab/recordings/word-open-save'
dotnet run --project src/OfficeHost.Tests --no-build -- --treenode-filter '/*/*/ReplayTests/Recording'
```

The report goes to the console and to `replay.txt` in the recording's folder.


## Recordings

A recording is a folder of exchanges in raw HTTP/1.1, numbered in the order the server received them:

- `0001.request.http`: the request line, the headers, a blank line, and the body as sent, without chunking, to the end
  of the file. Two headers of the recorder:
  - `X-Replay-User`: the user who signed in, as the server names them (`LAB\alice`, `i:0#.w|lab\alice`);
  - `X-Replay-Time`: when the server received the request, in ISO 8601.
- `0001.response.http`: the status line, the headers, a blank line, and the body, decoded.
- `files/`, optional: the documents as they were before the recording, by name.

The recordings hold the documents and the users' names: keep those of a real SharePoint out of the repository.
`TestServerHost.Record` records an OfficeHost to the same format.


## Recording the lab

The module of the shadow, [OfficeHost.SharePoint](../src/OfficeHost.SharePoint), records the lab's SharePoint from
inside it when its sink is a folder ([shadow.md](shadow.md)). It runs after the sign-in, so NTLM is not in its way.
- **What it records:** the cell storage and shared access requests of the web application, with their responses. It
  leaves out the 401s of the NTLM handshake, and copies an allowlist of request headers: never credentials or cookies.
- **How:** it takes the request's headers, user and time when the request is authenticated. The response goes through
  a filter that keeps a copy. At the response's first write, it reads the request body again from its start
  (SharePoint's handler read it through the buffered input stream: SharePoint SE serves `cellstorage.svc` with
  `/_vti_bin/cobalt.ashx/cellstorageservice`, as its IIS log shows), and a thread of its own writes the exchange to
  `C:\ReplayRecording` in the VM. So it reads nothing of cell storage before SharePoint does: reading the body first, an
  earlier recorder made PowerPoint drop its first request (see below). IsOnlyClient is read at authentication instead:
  after its ASMX handler (`sharedaccess.asmx`), even at the response's first write, the body reads empty, as the first
  versions of the module found in the lab. The earlier recorder read it so, and SharePoint answered all 27 as usual.
- **Errors** of the module go to the event log (source `OfficeHost.Shadow`), never to the request; exchanges it could not
  record, to `tombstones.txt` in the folder.

In an elevated PowerShell, from the repository (each script asks for the lab Administrator's password):

1. **Install it:** `src/sharepoint-lab/Install-ReplayRecorder.ps1`. It builds the module (Release), copies it and its
   install script into the VM, and runs the script there: a farm solution deployed to the GAC, then the module in the
   web.config ([shadow.md](shadow.md), "Install"). `-Sink` and `-ApiKey` send to a shadow instead of the folder. It
   removes the recorder of earlier versions of the lab first (a module in the bin folder).
2. **Record:** open, edit and save documents of `http://spse.mshome.net` in Office.
3. **Save it:** close Office, then `src/sharepoint-lab/Save-ReplayRecording.ps1 -Destination <folder>`. It copies the
   recording to the host and empties the folder in the VM for the next one.
4. **Replay it:** `OFFICEHOST_REPLAY` set to that folder, as above.
5. **Remove it:** `src/sharepoint-lab/Install-ReplayRecorder.ps1 -Remove`.

Office keeps a cache of the documents it opened. A document opened from the cache is queried with Office's
knowledge, and SharePoint answers only the changes, so the replay cannot start the document from that answer. Put the
document in `files/`, or clear the cache (File > Options > Save > Delete cached files).

**Tested** in the lab (SharePoint Server SE 16.0.14326) with the recorder before the module (the same capture, from a
bin folder): it recorded Word's open, polls, save and exit, and SharePoint answered as without it. IsOnlyClient is
recorded without a user, as SharePoint serves `sharedaccess.asmx` without sign-in; the replay sends it as alice.


## What it replays

- **Requests:** the cell storage service (`_vti_bin/cellstorage.svc/CellStorageService`) and IsOnlyClient
  (`_vti_bin/sharedaccess.asmx`), in order. The others (OPTIONS, GET, WebDAV, FrontPage RPC) are counted, not replayed.
- **Users:** each user of the recording is a user of OfficeHost, with the name without its domain or claim prefix
  (`LAB\alice` is `alice`).
- **Time:** the clock of OfficeHost is set to each request's `X-Replay-Time`, so that the coauthoring timeouts run as
  they did.
- **Documents:** each document a Request Url names starts as its file in `files/`. Without one, it starts as the file
  of the first content query that SharePoint answered in full. A document SharePoint did not have when the recording
  began, as one made by Save As, is named through the site of the last one, as Office does.
- **Ids:** SharePoint's ResourceID and TransitionID are learned from each pair of answers. They are replaced by
  OfficeHost's in the requests after that, for example in the `id` of IsOnlyClient and the `TransitionID` of AmIAlone.


## What it compares

For each SubResponse, by its SubRequestToken:
- the `ErrorCode` and `ErrorMessage`;
- the attributes of the SubResponseData. Those each server makes its own (`TransitionID`, `Etag`, the times, the user
  names, `Url`) are compared by whether they are there, not by value;
- the keys of its lists, such as the properties of GetDocMetaInfo: those only one server has;
- for a Cell, its FSSHTTPB sub-responses with their errors (`PutChanges Cell 12`, or `unreadable` with the codec's
  error), and the hash of the file that a query answers in full. When the two files differ, both are written to
  `differences/` in the recording's folder.
- for the editors table (a query of its partition), the editors by their CacheID, permission and metadata keys: not
  their login names (the users are mapped), their timeouts (OfficeHost keeps clients an hour at least), nor the layout
  of the XML.
- for a translated Cell, the copies of the two servers after it rather than its answers (see "The cells").

IsOnlyClient is compared by its result, and a request either server failed by its HTTP status.


## The cells

Office sends back the ids SharePoint gave the cell: its Storage Index, its knowledge, and the data elements of the
file. OfficeHost made other ids when it chunked the same file, so a Put Changes, or a query with SharePoint's
knowledge, refers to ids OfficeHost does not have. `ReplayCells` translates them. For each document and partition it
keeps:
- **the client's copy:** what SharePoint sent its clients, and what they put;
- **OfficeHost's copy:** what OfficeHost sent the replay, and what the replay put;
- **the knowledge of each server, in pairs:** from the answers both gave to the same Cell.

Then:
- **A query** gets OfficeHost's knowledge of the pair for SharePoint's, whichever client sends it, or else the latest
  OfficeHost gave.
- **A put of the content** is made again: the file the client put is read from the client's copy with its data
  elements, and put to OfficeHost against OfficeHost's copy, as `WordClient.Save` puts: new data elements, a new
  Storage Index, and OfficeHost's as the expected one. So OfficeHost is tested with a put of the replay's making, of
  the file Office put, not with Office's bytes.
- **A put of another partition** (Office's metadata), or of a whole file expecting nothing (a Save As), has the
  client's ids only, and goes as it was.
- **After each Cell,** the two copies are compared rather than the two answers: the file (its hash), or the editors of
  the editors table. Each answer holds only what was new to its client, so the copies are what the clients have.

These Cells are labelled `(translated)`. One that cannot be translated is labelled `(Office's cache)`: a client that
opened the document from Office's cache before the recording has data elements the recording never had, so its
knowledge is not known to the replay (OfficeHost answers with the whole file), and the file of its put cannot be read
(it goes as it was, and OfficeHost answers a coherency failure, `PutChanges Cell 12`). To test the puts, record
documents Office has not opened: new ones, uploaded for the recording.

The lock and coauthoring sub-requests do not depend on the cells, and are compared exactly all through a recording.
`ReplayTests.OfficeHostAgainstItself` replays a session of two Word clients recorded with OfficeHost to another: the
same answers, the save and the query after it translated, and the same file after the four Cells of the content.
`ReplayTests.OfficeHostFromCache` shows the labels of a client that knew the document before the recording.

## The first recording

On 2026-10-28, Word 16 opened `Lab.docx` of the lab as alice through COM, added a paragraph, saved and closed: 8
exchanges (the open, four IsOnlyClient polls, the save and its GetDocMetaInfo, the exit). The first replay gave 14 the
same and 7 different. Three of the differences were OfficeHost's, now answered as SharePoint does
(docs/fsshttp-notes.md, "Word's session replayed"): the Query Access answer the codec could not read, the
`JoinCoauthTime` of the editors table, and `HaveOnlyDemotionChanges` with the properties of a save. Now 16 are the same,
and the 5 left are known: GetVersions, the `vti_` properties of GetDocMetaInfo (twice), and the two untranslated Cells.

A second recording, the same day, `office-scenarios`: 47 exchanges of four sessions, each on its own document.
- **Word coauthoring** `Lab.docx` with a second client: two Word instances of one Windows user cannot open the same
  document, so the second client is a script sending Word's Coauth sub-requests (join, status, refresh, exit) as
  alice, with a client id of its own. SharePoint answered Word `Coauthoring` until it left.
- **A Save As** of `Probe.docx` to a new name, and a save of the copy.
- **Excel** on `Lab.xlsx`, which opened it read-only through COM and saved nothing.
- **PowerPoint** on `Lab.pptx`.

Its first replay found three more of OfficeHost's (docs/fsshttp-notes.md: the `TransitionID` of a refresh, the
`LockType` of a Save As, file properties for the content only) and refined `HaveOnlyDemotionChanges`. Now 93 are the
same, and the 35 left are of the known kinds. It also found two of the replay's: a document made by a Save As in the
recording is no longer seeded, and the files of the seeds come from content queries only.

With the cells translated, it gives 106 the same and 22 different: the saves of `Lab.docx` and `Lab.pptx`, and the
queries of documents opened from Office's cache, are `(Office's cache)`; the copies were compared after 13 Cells.

A third recording, `fresh-documents`: 23 exchanges on two documents uploaded for it, which Office had never opened.
Word saved twice, closed, opened the document again and saved; PowerPoint saved once. All four puts were translated and
applied as on SharePoint, and the copies held the same file after each of the 11 Cells compared: 49 the same, and of
the 10 different, 9 are GetVersions and GetDocMetaInfo. The tenth: when Word opened the document again, SharePoint
sent it changes (`HaveOnlyDemotionChanges="False"`) though Word had made every change itself, and the file of the
copies stayed the same: something of SharePoint's own in the cell, as a new revision, that OfficeHost does not have.

One request was recorded without a body, in both recordings with PowerPoint: its first, `GetMetadata_EnsureDocMetaInfo`
on a keep-alive connection, which declared a `Content-Length` (1043, 1015). SharePoint answered it with its error page
(a 302, "Object reference not set to an instance of an object") in 25 to 45 ms, and IIS logged `sc-win32-status` 64:
PowerPoint had dropped the connection. PowerPoint sent the same request again 30 ms later, on a new connection, which
succeeded. The IIS log of 2026-10-21, before the recorder, has 12 cell storage requests of PowerPoint answered 200 and no
302; that of 2026-10-28 has the two 302s, both with the recorder, which then read the request body at authentication.
The recorder now reads it after SharePoint's handler, and PowerPoint's next sessions (06:46, the first in a new
process) had none: 9 more requests, all 200. The replay reports such a request and goes on.
`src/sharepoint-lab/Find-CellStorageErrors.ps1` lists the cell storage requests of the lab's IIS logs that were
answered neither 200 nor 401, and with `-Around` the raw lines of a moment.

# Replay

The replay compares OfficeHost with SharePoint on real traffic. A recording of what desktop Office sent SharePoint, and
what SharePoint answered, is replayed to an OfficeHost of its own, request by request. Then each answer is compared
with SharePoint's. [Compare-CellStorage.ps1](sharepoint-lab.md) sends sub-requests written by hand; the replay sends
the ones Office sent.

The code is in [src/OfficeHost.Tests/Replay](../src/OfficeHost.Tests/Replay). `Replayer` replays, and `ReplayTests`
runs it:

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

[src/sharepoint-lab/ReplayRecorder.cs](../src/sharepoint-lab/ReplayRecorder.cs) is an ASP.NET module that records the
lab's SharePoint from inside it. It runs after the sign-in, so NTLM is not in its way.
- **What it records:** the cell storage and shared access requests of the web application, with their responses. It
  leaves out the 401s of the NTLM handshake, and the credentials (`Authorization`, cookies).
- **How:** it reads the request body when the request is authenticated and rewinds it for the service. WCF reads it
  again in the ASP.NET compatibility mode, which SharePoint enables. The response goes through a filter that keeps a
  copy, and the exchange is written when the response ends, to `C:\ReplayRecording` in the VM.
- **Errors** of the recorder go to `errors.txt` in that folder, never to the request.

In an elevated PowerShell, from the repository (each script asks for the lab Administrator's password):

1. **Install it:** `src/sharepoint-lab/Install-ReplayRecorder.ps1`. It compiles the module in the VM with the C#
   compiler of the .NET Framework (so the module is C# 5) into the bin folder of the web application. It then adds the
   module and its folder to the web.config as SharePoint web.config modifications. The bin folder needs the trust level
   `Full`, which the script checks.
2. **Record:** open, edit and save documents of `http://spse.mshome.net` in Office.
3. **Save it:** close Office, then `src/sharepoint-lab/Save-ReplayRecording.ps1 -Destination <folder>`. It copies the
   recording to the host and empties the folder in the VM for the next one.
4. **Replay it:** `OFFICEHOST_REPLAY` set to that folder, as above.
5. **Remove it:** `src/sharepoint-lab/Install-ReplayRecorder.ps1 -Remove`.

Office keeps a cache of the documents it opened. A document opened from the cache is queried with Office's
knowledge, and SharePoint answers only the changes, so the replay cannot start the document from that answer. Put the
document in `files/`, or clear the cache (File > Options > Save > Delete cached files).

**Tested** in the lab (SharePoint Server SE 16.0.14326): it recorded Word's open, polls, save and exit, and SharePoint
answered as without it. IsOnlyClient is recorded without a user, as SharePoint serves `sharedaccess.asmx` without
sign-in; the replay sends it as alice. The install checks what it did, and stops if the DLL, the web.config entry or the
module in IIS for `_vti_bin` is missing.


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

IsOnlyClient is compared by its result, and a request either server failed by its HTTP status.


## Not translated: the cells

The Cells are where the replay stops being exact. Office sends back the ids SharePoint gave the cell: its Storage
Index, its knowledge, and the data elements of the file. OfficeHost made other ids when it chunked the same file. So
a Put Changes, or a query with SharePoint's knowledge, refers to ids OfficeHost does not have. Their Cells are
labelled `(untranslated)`, and their differences are expected:

- **Put Changes:** OfficeHost answers a coherency failure (`CellRequestFail`, `PutChanges Cell 12`) where SharePoint
  applied it. From then on, OfficeHost's copy of the file stays the version before the save.
- **A query with knowledge:** OfficeHost does not know SharePoint's knowledge, so it answers with the whole file where
  SharePoint sent only the changes.

The lock and coauthoring sub-requests do not depend on the cells, and are compared exactly all through a recording.
`ReplayTests.OfficeHostAgainstItself` shows both: a session of two Word clients recorded with OfficeHost and replayed to
another gives the same answers, except for one Put Changes and the query after it.

## The first recording

On 2026-10-28, Word 16 opened `Lab.docx` of the lab as alice through COM, added a paragraph, saved and closed: 8
exchanges (the open, four IsOnlyClient polls, the save and its GetDocMetaInfo, the exit). The first replay gave 14 the
same and 7 different. Three of the differences were OfficeHost's, now answered as SharePoint does
(docs/fsshttp-notes.md, "Word's session replayed"): the Query Access answer the codec could not read, the
`JoinCoauthTime` of the editors table, and `HaveOnlyDemotionChanges` with the properties of a save. Now 16 are the same,
and the 5 left are known: GetVersions, the `vti_` properties of GetDocMetaInfo (twice), and the two untranslated Cells.

To translate the cells, the replay would keep a copy of each server's cell. It would read the file each Put Changes
writes against SharePoint's copy, and put the same file to OfficeHost against OfficeHost's copy, as `WordClient.Save`
does. Then it would compare the outcomes and the files. That tests OfficeHost with a Put Changes of the replay's
making, not of Office's.

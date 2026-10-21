# Desktop Office: manual checklist

The tests simulate Word (`src/OfficeHost.Tests/FssHttp`), but only real Word shows what Word does. This is the check of
the desktop site against Word by hand, before a release or after a change to `FssHttp/Service`. Record anything Word
does that the spec does not say, or that differs from what is expected here, in [fsshttp-notes.md](fsshttp-notes.md),
with the requests that show it.


## Setup

- **Certificate:** Word only uses FSSHTTP over HTTPS with a certificate it trusts. On the machine of the sample, trust
  the ASP.NET Core development certificate once (it asks before adding it to the Windows store):
  `dotnet dev-certs https --trust`.
- **Run:** `dotnet run --project src/OfficeHost.Sample --launch-profile desktop`. The profile listens on
  `https://localhost:5001` and logs every request and response to `src/OfficeHost.Sample/request-logs`
  (`{sequence}-{method}-{path}.txt`). The state of the desktop sessions is in `src/OfficeHost.Sample/fsshttp-state`,
  the files in `src/OfficeHost.Sample/files`.
- **Open:** `https://localhost:5001`, then "Open in desktop Office" for a `.docx`, or
  `https://localhost:5001/desktop/{fileId}?user=bob` to open it as another user.
- **A second Word:** a second Word on the same Windows account shares the first one's cache and identity, so it is not
  a second coauthor. Use either:
  - **another Windows account on the same machine** (fast user switching). It needs the development certificate in its
    own store: export it from the first account with
    `dotnet dev-certs https --export-path dev.crt --format PEM --no-password`, then in the second account
    `certutil -user -addstore Root dev.crt`;
  - **another machine.** It needs a certificate for the host name of the sample that it trusts (from your
    organisation's CA, for instance), given to Kestrel with `Kestrel__Certificates__Default__Path` and
    `Kestrel__Certificates__Default__Password`, `URLS=https://0.0.0.0:5001`, and
    `OfficeHost__DesktopBaseUrl=https://{host}:5001` so that the links name the host rather than `localhost`.
- **Before each run:** close Word, and clear `request-logs` to make the logs of the run easy to read.


## One user

| # | Do | Expect |
|---|---|---|
| 1 | Open a `.docx` from the sample. | Word opens it. On `localhost` (the Internet zone) it is in Protected View, after a join and an `ExitCoauthoring` (`Scenario: ReleaseLock_LockRelease`). |
| 2 | Enable Editing. | A second open: Coauth `JoinCoauthoring` answers `Alone`, the three Cell queries (editors table, metadata, content) succeed, then `IsOnlyClient` polls of `sharedaccess.asmx`. |
| 3 | Type, then save (Ctrl+S). | One Cell `PutChanges` with `BypassLockID`, answered with a Put Changes Response (no coherency failure). The file in `files` has the edit: open it in another application to check. |
| 4 | Edit and save again. | The second save also succeeds, and uploads only the changed chunks (a much smaller request than the file). |
| 5 | Close Word. | Coauth `ExitCoauthoring` and SchemaLock `ReleaseLock`; nothing in the `IOfficeLockStore` (a WOPI client can now lock the file). |
| 6 | Open the file again. | The content Cell query answers with what Word lacks; Word opens without a full download when its cache has the file. |
| 7 | Change the file outside Word (replace it in `files`) while it is closed, and open it. | Word shows the new content. |
| 8 | Change the file outside Word while Word has it open, then save in Word. | The save fails coherency, Word queries, gets the change, and saves again (or tells the user of a conflict). The file is never left with only one of the two edits without Word saying so. |
| 9 | Stop the sample with Word open for a minute, start it again, then save. | Word goes on polling without joining again, and saves: the state is in `fsshttp-state`, and the first poll writes the lock again. |
| 10 | Leave Word open and idle for an hour, then save. | `IsOnlyClient` every 30 seconds, `RefreshCoauthoring` 40 minutes after the join (the join's `Timeout` is 2580 seconds), and the save succeeds. |


## WOPI and desktop Office

Needs a WOPI client, as for the editor tests of the readme (Collabora Online in Docker).

| # | Do | Expect |
|---|---|---|
| 11 | Open a file in Word, then "Edit" in the sample. | The WOPI client cannot lock it: it opens read-only or says the file is locked. |
| 12 | Close Word, then "Edit". | The WOPI client edits and saves. |
| 13 | Close the WOPI client, then open the file in Word. | Word has the WOPI client's edits. |
| 14 | Edit in the WOPI client, then open the file in Word. | Word cannot join (`FileAlreadyLockedOnServer`), and opens read-only or says the file is locked. |


## Two users coauthoring

Alice is the first Word, Bob the second (see Setup), on the same file.

| # | Do | Expect |
|---|---|---|
| 15 | Alice opens and enables editing; then Bob does. | Bob's join answers `Coauthoring` with Alice's `TransitionID`. Alice's `IsOnlyClient` turns false, she sends `MarkTransitionComplete`, and both then poll Coauth `GetCoauthoringStatus` with the three Cell queries, about every 8 seconds. |
| 16 | Look at who is editing. | Each sees the other in Word's list of the people editing, with the name of the token (`?user=`). |
| 17 | Alice edits a paragraph and saves. | Bob's next poll brings the changes; Word shows them as updates available, and merges them at Bob's next save. |
| 18 | Both edit different paragraphs, then both save. | Both edits are in the file and in both Words. A Put Changes may fail coherency once; Word then queries, merges and saves. |
| 19 | Both edit the same sentence, then both save. | Word shows the conflict to the second to save, and the file ends up with the version they keep. |
| 20 | Bob closes. | Bob exits; Alice's status turns `Alone` and she goes back to `IsOnlyClient` polls. |
| 21 | Bob's Word is killed (Task Manager) instead of closed. | Alice stays in the session. Within an hour of Bob's last request (his `Timeout`), Alice is alone again. |
| 22 | Alice closes. | The lock is released. |


## Excel

Put an `.xlsx` in `files`. Excel does not coauthor: it takes an exclusive lock when the user first types. Edit by typing
in Excel: an edit through its object model does not take the lock.

| # | Do | Expect |
|---|---|---|
| 23 | Open the `.xlsx` from the sample, and Enable Editing. | `WhoAmI`, `GetDocMetaInfo` (with `vti_canmaybeedit=true`), then `OpenFile_Open`: a Cell query, `GetVersions` (`RequestNotSupported`) and a Query Access. No lock yet. |
| 24 | Type in a cell. | ExclusiveLock `GetLock` (`Timeout="600"`), answered `Success`, with its `LockStatus` answered `DependentOnlyOnFailRequestSucceeded`. |
| 25 | Save (Ctrl+S). | A Cell `PutChanges` with `BypassLockID`, answered with a Put Changes Response; the file in `files` has the edit. |
| 26 | Close Excel. | ExclusiveLock `ReleaseLock`; nothing in the `IOfficeLockStore`. |
| 27 | Open the file as Alice and type, then open it as Bob (`?user=bob`, see Setup) and type. | Bob's `GetLock` fails `FileAlreadyLockedOnServer` with `ErrorMessage="Alice …"`, and his `LockStatus` has `LockType="2"`. Excel shows File In Use: "locked for editing by Alice …". |
| 28 | Bob keeps "Receive a notification" and clicks OK; Alice closes Excel. | Bob's Excel polls `CheckLockAvailability` about every 30 seconds, and shows File Now Available once Alice's lock is gone. Its Read-Write stays read-only, as against SharePoint: Bob closes and reopens the workbook to edit it. |


## Other applications

PowerPoint coauthors as Word does. Put a `.pptx` in `files`.

| # | Do | Expect |
|---|---|---|
| 29 | Open the `.pptx` from the sample, and Enable Editing. | `WhoAmI` and `GetDocMetaInfo`, then `JoinCoauthoring` (with `ExclusiveLockID` and `AllowFallbackToExclusive`) answered `Alone`, the editors table and content queries, then `IsOnlyClient` polls. |
| 30 | Edit a slide, save, close. | `AutoSaveFile_Save` answered with a Put Changes Response, the file in `files` has the edit, and `ExitCoauthoring` releases the lock. |

Other applications (Visio, Project, OneNote) get no links from `LaunchDesktop`, nor do the legacy binary formats
(`.doc`, `.xls`, `.ppt` and their templates) and `.rtf`, which are not supported.


## Last run

- **Word** 16.0.20326 on Windows 11, one Windows account: steps 1, 2, 3, 5, 9 and 10 passed (2026-10-21).
- **Excel** 16.0.20326, one Windows account: steps 23 to 26 passed (2026-10-21), and 27 and 28 with the other user's
  lock taken by a script as user bob; the lock-out also against SharePoint Server with two users (see the notes).
- **PowerPoint** 16.0.20326: steps 29 and 30 passed (2026-10-21), editing through COM.
- What they sent is in [fsshttp-notes.md](fsshttp-notes.md).
- **Not yet run:** the WOPI client steps, Word coauthoring with two users, and Excel against Excel, which need a second
  Windows account or machine.

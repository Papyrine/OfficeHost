# SharePoint lab

A SharePoint Server of our own, to capture what desktop Office sends SharePoint and compare it with OfficeHost: one
Hyper-V VM with a domain, SQL Server 2022 Developer and SharePoint Server Subscription Edition, from Microsoft's free
evaluations. It has test users of its own (`LAB\alice`, `LAB\bob`), so captures from it hold no one's data and can be
shared, and a second user is a second sign-in, not a second person.

The scripts are in [src/sharepoint-lab](../src/sharepoint-lab). Each step is a script, and each script can be run
again: it goes on from where it stopped.


## What it needs

- **Hyper-V,** and an elevated PowerShell: building the VM disk mounts a VHDX, which needs an administrator.
- **Disk:** about 80 GB free on the drive of the VM. The media are 7.3 GB (the Windows ISO, 4.7 GB of it, can go once
  the VM is built). The VM's disk is dynamic: it takes what is written, about 55 to 65 GB once installed (Windows 12 to
  15 GB, its page file up to 16 GB, SQL Server 2 GB, SharePoint 8 to 10 GB, the databases 3 to 5 GB), and can grow to
  the `-DiskGB` of `New-LabVm.ps1`, 150 GB by default; `-DiskGB 80` caps it lower.
- **Memory:** the VM has dynamic memory: it starts with 8 GB and takes up to 16 GB (SharePoint's minimum for a single
  server farm) when the host has it, so 32 GB on the host. Stop Docker Desktop for a faster farm.
- **Internet** in the VM, through the Default Switch: the Windows evaluation activates online in its first 10 days.

The evaluations last 180 days. Licensing: evaluation only.


## Media

`Get-LabMedia.ps1` downloads them, from Microsoft only, and needs no elevation:

| File | Size | Source |
|---|---|---|
| `WindowsServer2022.iso` | 4.7 GB | Windows Server 2022 evaluation, Evaluation Center |
| `OfficeServer.iso` | 1.4 GB | SharePoint Server Subscription Edition, [Download Center 103599](https://www.microsoft.com/en-us/download/details.aspx?id=103599) |
| `sql\SQLServer2022-x64-ENU-Dev.iso` | 1.1 GB | SQL Server 2022 Developer, download.microsoft.com (the media of its bootstrapper, which crashes on a machine with SQL Server 2025) |

```bash
pwsh src/sharepoint-lab/Get-LabMedia.ps1 -Path D:/SharePointLab/media
```

The trial product key of SharePoint is on the Download Center page, under Install Instructions: the Enterprise one.


## Build

In an elevated PowerShell, from the repository:

1. **The VM** (about 15 minutes): Windows Server 2022 Datacenter applied to a new VHDX, a first boot without
   questions, the SQL Server and SharePoint ISOs in two DVD drives. It asks for the lab Administrator's password.

   ```bash
   src/sharepoint-lab/New-LabVm.ps1 -Media D:/SharePointLab/media -Path D:/SharePointLab
   ```
2. **The farm** (one to two hours, with several restarts): the domain `lab.test` (`LAB`) with the VM as its domain
   controller, the users `alice` and `bob`, SQL Server, the SharePoint prerequisites and binaries, the farm with
   Central Administration on port 2016, and a team site where alice and bob can edit. It asks for the Administrator's
   password again, and a password for the test users.

   ```bash
   src/sharepoint-lab/Install-LabFarm.ps1 -ProductKey <the Enterprise trial key>
   ```
3. **A checkpoint,** to come back to a clean farm: `Checkpoint-VM -Name SPSE -SnapshotName Installed`.

`LAB\Administrator` is the farm account and the application pool's: fine for a lab, not a farm to copy. The domain
controller, SQL Server and SharePoint share the VM; SharePoint SE on a domain controller is for development only. If
the SharePoint setup refuses it, the domain goes to a second, small VM.


## Use

- **The site:** `http://spse.mshome.net` (the Default Switch names the VM `<name>.mshome.net` for the host), its
  documents at `http://spse.mshome.net/Shared%20Documents`. Upload a `.docx` and an `.xlsx` in the browser, signed in as
  `LAB\alice`.
- **Office:** "Open in app" from the library, or `ms-word:ofe|u|http://spse.mshome.net/Shared%20Documents/<file>`.
  Office asks for credentials the first time (the host is not in the domain): `LAB\alice` and her password.
- **Capture:** `pwsh src/fsshttp-capture.ps1 -HostName spse.mshome.net -Seconds 300`, then open the document.
- **A second user:** Office keeps one sign-in per host, so bob is a second Windows account on the host, another
  machine, or a script sending bob's requests with his credentials (`-Credential` of `Invoke-WebRequest`).
- **After a restart of the host,** the Default Switch hands out other addresses: start the VM, and the names follow.


## Status

Built on 2026-10-21 on Windows 11 (32 GB, Docker running): SharePoint answers at `http://spse.mshome.net` with NTLM,
`MicrosoftSharePointTeamServices: 16.0.0.14326`. The first run needed these, now in the scripts:

- **bcdboot** of the host failed on the image (193): the image's own bcdboot.
- **Memory:** 16 GB of fixed memory was not free beside Docker: dynamic memory, 8 to 16 GB.
- **Passwords:** the domain's recovery password and the farm passphrase are the Administrator's, so it has to be
  complex: checked before anything starts.
- **The domain controller:** its DNS pointed at the Default Switch, and Active Directory Web Services was left
  disabled: the VM's DNS is itself (forwarding to the switch), and ADWS is set to automatic.
- **SharePoint setup:** 30203 without `/IAcceptTheLicenseTerms` on the command line.
- **The farm:** created before SQL Server took connections after the restart: it waits for SQL Server.
  `Install-SPService` timed out once: the configuration after the farm runs again on each run.

**First capture** (Word and Excel 16.0.20326, signed in as `LAB\alice`, which Office asked for once each): the same as
the captures of docs/fsshttp-notes.md. Word joins coauthoring (`Alone`), exits and joins again around Protected View,
polls `IsOnlyClient`, saves, and exits; Excel opens with no lock, takes an exclusive lock at the first edit, saves, and
releases it. Word's `IsOnlyClient` polls are the same as against OfficeHost: a burst every 8 seconds after the open,
then every 30 seconds (docs/fsshttp-notes.md).

To upload files without a password in a script: sign in to the site in a browser, and use the classic upload page,
`/_layouts/15/Upload.aspx?List={id}` (the id from `/_api/web/lists/getbytitle('Documents')/Id`), whose file input
takes a file directly.

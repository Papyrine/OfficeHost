/// <summary>
/// The storage interface tests of the Python implementation (test/test_storageiface.py), against the local storage.
/// </summary>
public class StorageTests : IDisposable
{
    const string databuf = "ebe5tresbsrdthbrdhvdtr";
    const string endpoint = "/";
    const string userId = "0:0";

    readonly TempFolder folder = new();
    readonly ManualTime time = new();
    readonly LocalStorage storage;

    public StorageTests()
    {
        var config = Helpers.Ini(
            $"""
             [general]
             storagetype = local
             wopilockexpiration = 3
             [io]
             chunksize = 1
             [local]
             storagehomepath = {folder.Path}
             """);
        storage = new(new(config), NullLogger.Instance, time);
    }

    public void Dispose() =>
        folder.Dispose();

    Task Write(string path, string content, LockMetadata? lockMetadata = null, bool isLock = false, bool noVersion = false) =>
        storage.WriteFile(endpoint, path, userId, Helpers.Stream(content), Encoding.UTF8.GetByteCount(content), lockMetadata, isLock, noVersion);

    async Task<string> Read(string path)
    {
        await using var stream = await storage.ReadFile(endpoint, path, userId, null);
        using var reader = new StreamReader(stream);
        return await reader.ReadToEndAsync();
    }

    [Test]
    public async Task Stat()
    {
        await Write("/test.txt", databuf);
        var statInfo = await storage.Stat(endpoint, "/test.txt", userId);
        await Assert.That(statInfo.Size).IsEqualTo(databuf.Length);
        await Assert.That(statInfo.MTime).IsGreaterThan(0);
        await storage.RemoveFile(endpoint, "/test.txt", userId);
    }

    [Test]
    public async Task StatX_FileId()
    {
        await Write("/test.txt", databuf);
        var statInfo = await storage.StatX(endpoint, "/test.txt", userId);
        await Assert.That(statInfo.Inode).StartsWith("local!");
        await Assert.That(statInfo.FilePath).IsEqualTo("/test.txt");
        await Assert.That(statInfo.OwnerId).IsEqualTo("0:0");
        await Assert.That(statInfo.Etag).IsNotEmpty();
        await Assert.That(statInfo.Xattrs).IsEmpty();
        await storage.RemoveFile(endpoint, "/test.txt", userId);
    }

    [Test]
    public async Task StatX_InvariantFileId()
    {
        await Write("/test&upd.txt", databuf);
        var inode = (await storage.StatX(endpoint, "/test&upd.txt", userId)).Inode;
        await Write("/test&upd.txt", databuf);
        var statInfo = await storage.StatX(endpoint, "/test&upd.txt", userId);
        await Assert.That(statInfo.Inode).IsEqualTo(inode);
        await storage.RemoveFile(endpoint, "/test&upd.txt", userId);
    }

    [Test]
    public async Task Stat_NoFile()
    {
        var exception = await Helpers.Throws<StorageException>(() => storage.Stat(endpoint, "/hopefullynotexisting", userId));
        await Assert.That(exception.Message).Contains(Common.EnoentMsg);
    }

    [Test]
    public async Task StatX_NoFile()
    {
        var exception = await Helpers.Throws<StorageException>(() => storage.StatX(endpoint, "/hopefullynotexisting", userId));
        await Assert.That(exception.Message).Contains(Common.EnoentMsg);
    }

    [Test]
    public async Task Stat_Directory()
    {
        var exception = await Helpers.Throws<StorageException>(() => storage.Stat(endpoint, "/", userId));
        await Assert.That(exception.Message).IsEqualTo("Is a directory");
        await Assert.That(await storage.HealthCheck()).IsEqualTo("Warning");
    }

    [Test]
    public async Task ReadFile_Binary()
    {
        await Write("/test.txt", databuf);
        await Assert.That(await Read("/test.txt")).IsEqualTo(databuf);
        await storage.RemoveFile(endpoint, "/test.txt", userId);
    }

    [Test]
    public async Task ReadFile_Text()
    {
        await Write("/test.txt", "bla\n");
        await Assert.That(await Read("/test.txt")).IsEqualTo("bla\n");
        await storage.RemoveFile(endpoint, "/test.txt", userId);
    }

    [Test]
    public async Task ReadFile_Empty()
    {
        await Write("/test.txt", "");
        await Assert.That(await Read("/test.txt")).IsEqualTo("");
        await storage.RemoveFile(endpoint, "/test.txt", userId);
    }

    [Test]
    public async Task Read_NoFile()
    {
        var exception = await Helpers.Throws<StorageException>(() => storage.ReadFile(endpoint, "/hopefullynotexisting", userId, null));
        await Assert.That(exception.Message).Contains(Common.EnoentMsg);
    }

    [Test]
    public async Task Write_Remove_SpecialChars()
    {
        await Write("/testwrite&rm", databuf);
        await storage.Stat(endpoint, "/testwrite&rm", userId);
        await storage.RemoveFile(endpoint, "/testwrite&rm", userId);
        await Helpers.Throws<StorageException>(() => storage.Stat(endpoint, "/testwrite&rm", userId));
    }

    [Test]
    public async Task Write_SizeMismatch()
    {
        var exception = await Helpers.Throws<StorageException>(() => storage.WriteFile(endpoint, "/test.txt", userId, Helpers.Stream("abc"), 4, null));
        await Assert.That(exception.Message).IsEqualTo("Written 3 bytes but content is 4 bytes");
    }

    [Test]
    public async Task Write_NoVersion()
    {
        await Write("/test.txt", "bla\n");
        await Write("/test.txt", "blabla\n", noVersion: true);
        await Assert.That(await Read("/test.txt")).IsEqualTo("blabla\n");
        await storage.RemoveFile(endpoint, "/test.txt", userId);
    }

    [Test]
    public async Task Write_IsLock()
    {
        await Write("/testoverwrite", databuf, isLock: true);
        await storage.Stat(endpoint, "/testoverwrite", userId);
        var exception = await Helpers.Throws<StorageException>(() => Write("/testoverwrite", databuf, isLock: true));
        await Assert.That(exception.Message).Contains(Common.ExclError);
        await storage.RemoveFile(endpoint, "/testoverwrite", userId);
    }

    [Test]
    public async Task Write_Race()
    {
        var results = await Task.WhenAll(
            Enumerable.Range(0, 2)
                .Select(_ => Capture(() => Write("/testwriterace", databuf, isLock: true))));
        await Assert.That(results.Count(_ => _ == null)).IsEqualTo(1);
        await Assert.That(results.Single(_ => _ != null)!.Message).Contains(Common.ExclError);
        await storage.RemoveFile(endpoint, "/testwriterace", userId);
    }

    static async Task<StorageException?> Capture(Func<Task> action)
    {
        await Task.Yield();
        try
        {
            await action();
            return null;
        }
        catch (StorageException exception)
        {
            return exception;
        }
    }

    [Test]
    public async Task Lock()
    {
        await Write("/testlock", databuf);
        await storage.SetLock(endpoint, "/testlock", userId, "test app", WopiUtils.EncodeLock("testlock"));
        var storageLock = await storage.GetLock(endpoint, "/testlock", userId);
        await Assert.That(storageLock).IsNotNull();
        await Assert.That(storageLock!.LockId).IsEqualTo(WopiUtils.EncodeLock("testlock"));
        await Assert.That(storageLock.AppName).IsEqualTo("test app");
        await Assert.That(storageLock.Expiration.Seconds).IsEqualTo(time.Now.ToUnixTimeSeconds() + 3);
        var exception = await Helpers.Throws<StorageException>(() => storage.SetLock(endpoint, "/testlock", userId, "test app", "mismatchlock"));
        await Assert.That(exception.Message).Contains(Common.ExclError);
        await storage.Unlock(endpoint, "/testlock", userId, "test app", WopiUtils.EncodeLock("testlock"));
        await Assert.That(await storage.GetLock(endpoint, "/testlock", userId)).IsNull();
        await storage.RemoveFile(endpoint, "/testlock", userId, force: true);
    }

    [Test]
    public async Task RefreshLock()
    {
        await Write("/testrlock", databuf);
        var exception = await Helpers.Throws<StorageException>(() => storage.RefreshLock(endpoint, "/testrlock", userId, "test app", "testlock"));
        await Assert.That(exception.Message).IsEqualTo(Common.ExclError);
        await storage.SetLock(endpoint, "/testrlock", userId, "test app", "testlock");
        exception = await Helpers.Throws<StorageException>(() => storage.RefreshLock(endpoint, "/testrlock", userId, "test app", "newlock", "mismatch"));
        await Assert.That(exception.Message).IsEqualTo(Common.ExclError);
        await storage.RefreshLock(endpoint, "/testrlock", userId, "test app", "newlock", "testlock");
        var storageLock = await storage.GetLock(endpoint, "/testrlock", userId);
        await Assert.That(storageLock!.LockId).IsEqualTo("newlock");
        await Assert.That(storageLock.AppName).IsEqualTo("test app");
        await storage.RefreshLock(endpoint, "/testrlock", userId, "test app", "newlock");
        exception = await Helpers.Throws<StorageException>(() => storage.RefreshLock(endpoint, "/testrlock", userId, "mismatched app", "newlock"));
        await Assert.That(exception.Message).Contains(Common.ExclError);
        await storage.RemoveFile(endpoint, "/testrlock", userId);
    }

    [Test]
    public async Task Lock_Race()
    {
        await Write("/testlockrace", databuf);
        var results = await Task.WhenAll(
            Capture(() => storage.SetLock(endpoint, "/testlockrace", userId, "test app", "testlock")),
            Capture(() => storage.SetLock(endpoint, "/testlockrace", userId, "test app 2", "testlock2")));
        await Assert.That(results.Count(_ => _ == null)).IsEqualTo(1);
        await Assert.That(results.Single(_ => _ != null)!.Message).Contains(Common.ExclError);
        await storage.RemoveFile(endpoint, "/testlockrace", userId);
    }

    [Test]
    public async Task Lock_Operations()
    {
        await Write("/testlockop", databuf);
        const string payload = "{'foo': 'bar', 'xyz': 123, 'data': {'a': 'b'}}";
        var lockMetadata = new LockMetadata("test app", WopiUtils.EncodeLock(payload));
        await storage.SetLock(endpoint, "/testlockop", userId, "test app", WopiUtils.EncodeLock(payload));
        await Write("/testlockop", databuf, lockMetadata);
        // different interfaces raise exceptions on either mismatching app xor mismatching lock payload,
        // this is why we test that both mismatch
        await Helpers.Throws<StorageException>(() => Write("/testlockop", databuf, new("mismatch app", "mismatchlock")));
        // with xattrs, it's fine to set them without lock context (the lock should apply to files' contents only)
        await storage.SetXattr(endpoint, "/testlockop", userId, "testkey", "123", new("mismatch app", "mismatchlock"));
        await storage.SetXattr(endpoint, "/testlockop", userId, "testkey", "123", null);
        await storage.RemoveXattr(endpoint, "/testlockop", userId, "testkey", null);
        await storage.RefreshLock(endpoint, "/testlockop", userId, "test app", WopiUtils.EncodeLock(payload));
        await Helpers.Throws<StorageException>(() => storage.RefreshLock(endpoint, "/testlockop", userId, "mismatched app", "mismatchlock"));
        await Assert.That(await Read("/testlockop")).IsEqualTo(databuf);
        await Helpers.Throws<StorageException>(() => Write("/testlockop", databuf));
        await storage.RenameFile(endpoint, "/testlockop", "/testlockop_renamed", userId, lockMetadata);
        // the lock moved along with the file
        await Assert.That((await storage.GetLock(endpoint, "/testlockop_renamed", userId))!.AppName).IsEqualTo("test app");
        await storage.RemoveFile(endpoint, "/testlockop_renamed", userId);
    }

    [Test]
    public async Task Expired_Locks()
    {
        await Write("/testelock", databuf);
        // 1: get an expired lock
        await storage.SetLock(endpoint, "/testelock", userId, "test app", "testlock1");
        time.Advance(TimeSpan.FromSeconds(3.1));
        await Assert.That(await storage.GetLock(endpoint, "/testelock", userId)).IsNull();
        // 2: set a new lock over an expired one
        await storage.SetLock(endpoint, "/testelock", userId, "test app", "testlock2");
        await Assert.That((await storage.GetLock(endpoint, "/testelock", userId))!.LockId).IsEqualTo("testlock2");
        // 3: refresh an expired lock, should fail
        time.Advance(TimeSpan.FromSeconds(3.1));
        var exception = await Helpers.Throws<StorageException>(() => storage.RefreshLock(endpoint, "/testelock", userId, "test app", "testlock2"));
        await Assert.That(exception.Message).IsEqualTo(Common.ExclError);
        // 4: unlock an expired lock, should fail
        exception = await Helpers.Throws<StorageException>(() => storage.Unlock(endpoint, "/testelock", userId, "test app", "testlock2"));
        await Assert.That(exception.Message).IsEqualTo(Common.ExclError);
        // 5: overwrite a file having an expired lock with a mismatched one or without lock metadata
        await Write("/testelock", databuf, new("mismatched app", "mismatched lock"));
        await Write("/testelock", databuf);
        await storage.RemoveFile(endpoint, "/testelock", userId);
    }

    [Test]
    public async Task Remove_NoFile()
    {
        var exception = await Helpers.Throws<StorageException>(() => storage.RemoveFile(endpoint, "/hopefullynotexisting", userId));
        await Assert.That(exception.Message).Contains(Common.EnoentMsg);
    }

    [Test]
    public async Task Xattr()
    {
        await Write("/test&xattr.txt", databuf);
        await storage.SetXattr(endpoint, "/test&xattr.txt", userId, "testkey", "123", null);
        await storage.SetLock(endpoint, "/test&xattr.txt", userId, "test app", "xattrlock");
        await storage.SetXattr(endpoint, "/test&xattr.txt", userId, "testkey", "123", new("test app", "xattrlock"));
        var statInfo = await storage.StatX(endpoint, "/test&xattr.txt", userId);
        await Assert.That(statInfo.Xattrs["testkey"]).IsEqualTo("123");
        await storage.RemoveXattr(endpoint, "/test&xattr.txt", userId, "testkey", new("test app", "xattrlock"));
        statInfo = await storage.StatX(endpoint, "/test&xattr.txt", userId);
        await Assert.That(statInfo.Xattrs.ContainsKey("testkey")).IsFalse();
        await storage.RemoveFile(endpoint, "/test&xattr.txt", userId);
    }

    [Test]
    public async Task Rename_StatX()
    {
        await Write("/test.txt", databuf);
        var statInfo = await storage.StatX(endpoint, "/test.txt", userId);
        var pathRef = statInfo.FilePath[..statInfo.FilePath.LastIndexOf('/')];
        await storage.RenameFile(endpoint, "/test.txt", "/test&ren.txt", userId, null);
        statInfo = await storage.StatX(endpoint, "/test&ren.txt", userId);
        await Assert.That(statInfo.FilePath).IsEqualTo($"{pathRef}/test&ren.txt");
        await storage.RenameFile(endpoint, "/test&ren.txt", "/test.txt", userId, null);
        statInfo = await storage.StatX(endpoint, "/test.txt", userId);
        await Assert.That(statInfo.FilePath).IsEqualTo($"{pathRef}/test.txt");
        await storage.RemoveFile(endpoint, "/test.txt", userId);
    }

    [Test]
    public async Task Path_OutsideStorage()
    {
        var exception = await Helpers.Throws<StorageException>(() => storage.Stat(endpoint, "/../outside.txt", userId));
        await Assert.That(exception.Message).IsEqualTo(Common.AccessError);
        exception = await Helpers.Throws<StorageException>(() => storage.Stat(endpoint, "/.wopimeta/x.json", userId));
        await Assert.That(exception.Message).IsEqualTo(Common.AccessError);
    }
}

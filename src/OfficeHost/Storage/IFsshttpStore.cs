namespace OfficeHost;

/// <summary>
/// Keeps the state of desktop Office coauthoring ([MS-FSSHTTP]): for each file, and each partition of it (the content,
/// the metadata, the editors table), the data elements Office synchronizes, as a value only OfficeHost reads. The file
/// itself stays in the <see cref="IOfficeStorage"/>; a lost state only costs Office a full download of the file. The
/// built-in <see cref="MemoryFsshttpStore"/> suits a single instance of the application, and
/// <see cref="FileSystemFsshttpStore"/> a single machine; a store shared by all instances is needed otherwise.
/// </summary>
public interface IFsshttpStore
{
    /// <summary>
    /// The state of the partition of the file, or null when it has none.
    /// </summary>
    Task<FsshttpState?> Get(string fileId, Guid partition, Cancel cancel = default);

    /// <summary>
    /// Atomically replaces the state of the partition of the file with <paramref name="value"/> when the version of its
    /// current state is <paramref name="comparand"/> (null when it has none), and returns whether it did.
    /// </summary>
    Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FsshttpState value, Cancel cancel = default);

    /// <summary>
    /// Forgets the states of all the partitions of the file. A WOPI DeleteFile calls it; an application that deletes a
    /// file itself calls it after.
    /// </summary>
    Task Delete(string fileId, Cancel cancel = default);
}

/// <param name="Version">The version OfficeHost gives the state, compared as is.</param>
/// <param name="Data">The state, opaque to the store. The store must not change it.</param>
public record FsshttpState(string Version, ReadOnlyMemory<byte> Data);

using System.Buffers;
using System.Collections;
using System.Reflection;
using System.Text;

delegate void Writing(ref FsshttpbWriter writer);

delegate T Reading<out T>(ref FsshttpbReader reader);

/// <summary>
/// Helpers of the FSSHTTPB tests: hex, and writing or reading a part of a message on its own.
/// </summary>
static class Fsshttpb
{
    public static byte[] Hex(string hex) => Convert.FromHexString(hex.Replace(" ", "").Replace("\r", "").Replace("\n", ""));

    public static string Hex(ReadOnlySpan<byte> bytes) => Convert.ToHexString(bytes);

    public static byte[] Bytes(Writing write)
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FsshttpbWriter(output);
        write(ref writer);
        return output.WrittenSpan.ToArray();
    }

    /// <summary>
    /// Reads a part on its own, requiring all the bytes to be read.
    /// </summary>
    public static T Read<T>(byte[] bytes, Reading<T> read)
    {
        var reader = new FsshttpbReader(bytes);
        var result = read(ref reader);
        reader.EnsureEnd("the test data");
        return result;
    }

    public static Guid Guid(string value) => new(value);

    public static object ParseRequest(byte[] bytes) => FsshttpbRequest.Read(bytes);

    public static object ParseResponse(byte[] bytes) => FsshttpbResponse.Read(bytes);

    public static object ParseSubRequest(byte[] bytes) => Read(bytes, (ref FsshttpbReader _) => SubRequest.Read(ref _));

    public static object ParseSubResponse(byte[] bytes) => Read(bytes, (ref FsshttpbReader _) => SubResponse.Read(ref _));

    public static object ParseElement(byte[] bytes) => Read(bytes, (ref FsshttpbReader _) => DataElement.Read(ref _));

    public static object ParseKnowledge(byte[] bytes) => Read(bytes, (ref FsshttpbReader _) => Knowledge.Read(ref _));

    public static object ParseError(byte[] bytes) => Read(bytes, (ref FsshttpbReader _) => ResponseError.Read(ref _));

    /// <summary>
    /// Writes a part of a message on its own.
    /// </summary>
    public static byte[] Write(object value) =>
        value switch
        {
            FsshttpbRequest request => request.ToArray(),
            FsshttpbResponse response => response.ToArray(),
            SubRequest subRequest => Bytes((ref FsshttpbWriter _) => subRequest.Write(ref _)),
            SubResponse subResponse => subResponse.ToArray(),
            DataElement element => Bytes((ref FsshttpbWriter _) => element.Write(ref _)),
            Knowledge knowledge => Bytes((ref FsshttpbWriter _) => knowledge.Write(ref _)),
            ResponseError error => Bytes((ref FsshttpbWriter _) => error.Write(ref _)),
            _ => throw new ArgumentException(value.GetType().Name)
        };

    /// <summary>
    /// Describes a model of the codec with all its fields, deeply, to compare models: records that hold lists or memory
    /// have no deep equality.
    /// </summary>
    public static string Describe(object? value)
    {
        var builder = new StringBuilder();
        Describe(builder, value);
        return builder.ToString();
    }

    static void Describe(StringBuilder builder, object? value)
    {
        switch (value)
        {
            case null:
                builder.Append("null");
                return;
            case string text:
                builder.Append('"').Append(text).Append('"');
                return;
            case ReadOnlyMemory<byte> memory:
                builder.Append("0x").Append(Hex(memory.Span));
                return;
            case Enum or System.Guid or ExtendedGuid or SerialNumber:
                builder.Append(value);
                return;
            case IEnumerable items:
                builder.Append('[');
                foreach (var item in items)
                {
                    Describe(builder, item);
                    builder.Append(", ");
                }

                builder.Append(']');
                return;
        }

        var type = value.GetType();
        if (type.IsPrimitive)
        {
            builder.Append(value);
            return;
        }

        builder.Append(type.Name).Append(" { ");
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            builder.Append(property.Name).Append(" = ");
            Describe(builder, property.GetValue(value));
            builder.Append(", ");
        }

        builder.Append('}');
    }
}

/// <summary>
/// Byte-exact FSSHTTPB, as published or captured:
/// - [MS-FSSHTTPB] section 4: the Query Changes request (4.1), the Query Changes sub-response (4.2), the data elements of
///   the Put Changes request (4.3) and the Put Changes response (4.4). The Revision Manifest of 4.3 is cut short in the
///   spec: it ends here after its first Object Group reference.
/// - [MS-FSSHTTP] 4.1.1 and 4.1.2: the base64 of the Cell sub-requests and sub-responses of a file open.
/// - Word 16.0.20326: the Cell sub-requests of its open, sent to the spike of the sample (docs/fsshttp-spike.md, runs 6
///   and 7). They are queries, with no document content.
/// </summary>
static class FsshttpbVectors
{
    public const string SpecQueryChangesRequest =
        """
        0C 00 0B 00 9C CF 29 F3 39 94 06 9B 06 02 00 00
        EE 02 00 00 AA 02 20 00 7E B8 31 E7 45 DD AA 44
        AB 80 0C 75 FB D1 53 0E 7A 02 08 00 C4 27 A1 0F
        77 01 16 02 06 00 03 05 00 8A 02 02 00 00 DA 02
        06 00 03 00 00 CA 02 08 00 08 00 80 03 84 00 41
        0B 01 AC 02 00 55 03 01
        """;

    public const string SpecQueryChangesSubResponse =
        """
        0E 02 06 00 03 05 00 FA 02 24 00 0C FD 98 0D A0
        FD 40 99 4D 93 0A 63 22 D7 68 91 36 00 84 00 26
        02 20 00 F6 35 7A 32 61 07 14 44 96 86 51 E9 00
        66 7A 4D A4 00 78 28 80 93 0A E2 55 FD A5 BC 90
        37 45 1C 9D 86 E9 49 00 1C F9 08 78 28 7F 6C F5
        1D AA 02 5A 43 90 37 45 1C 9D 86 E9 49 00 FC F8
        08 51 13 01 26 02 20 00 0E E9 76 3A 32 80 0C 4D
        B9 DD F3 C6 50 29 43 3E 4C 01 20 2A 0C 7F 6C F5
        1D AA 02 5A 43 90 37 45 1C 9D 86 E9 49 FC F8 08
        00 A5 13 01 41 07 01
        """;

    public const string SpecPutChangesResponse =
        """
        0C 00 0B 00 9D CF 29 F3 39 94 06 9B 16 03 02 00
        00 0E 02 06 00 03 0B 00 84 00 26 02 20 00 F6 35
        7A 32 61 07 14 44 96 86 51 E9 00 66 7A 4D A4 00
        78 24 22 92 69 92 46 AD 53 B3 94 89 C2 4F 5A CF
        A0 9A 00 E9 78 24 DD 6D 96 6D B9 52 AC 4C 94 89
        C2 4F 5A CF A0 9A 00 DF 51 13 01 26 02 20 00 13
        1F 09 10 82 C8 FB 40 98 86 65 33 F9 34 C2 1D 6C
        01 70 2D 0C F9 0B 41 37 6F D1 99 44 A6 C3 27 23
        2E DC A7 11 09 33 00 00 00 B5 13 01 41 07 01 8B
        01
        """;

    public const string SpecPutChangesRequestHeader =
        """
        0C 00 0B 00 9C CF 29 F3 39 94 06 9B 06 02 00 00
        EE 02 00 00 AA 02 20 00 7E B8 31 E7 45 DD AA 44
        AB 80 0C 75 FB D1 53 0E 7A 02 08 00 B4 27 E1 2E
        77 01 16 02 06 00 03 0B 00 D2 02 26 00 0C 8E 2E
        2E 05 D1 C0 86 48 9C 51 29 D6 61 71 4F 67 00 48
        0B 01 AC 02 00
        """;

    public const string SpecStorageManifest =
        """
        0C 56 0C 99 FA 30 D7 2C 12 88 42 B7 22 0A 12 5C FD A7 E5 80 47 AF 30 54 71 6E 9B 40 98 06 70 7E 81 8D C1 02 32
        00 00 00 00 00 00 00 05 60 20 94 33 B9 0E 1D 57 E9 41 AA D3 88 0D 92 D3 19 55 38 66 14 B9 FA DE 84 A3 AA 0D 4A
        A3 A8 52 0C 77 AC 70 73 0C B9 FA DE 84 A3 AA 0D 4A A3 A8 52 0C 77 AC 70 73 0C 65 46 2A 6F C8 42 C7 46 BA B4 E2
        8F DC E1 E3 2B 05
        """;

    public const string SpecCellManifest =
        """
        0C 58 60 0C 8E FC 0B 2C 04 9B 61 4C AB 49 48 45 E6 03 EC A0 80 47 AF 30 54 71 6E 9B 40 98 06 70 7E 81 8D C1 02
        33 00 00 00 00 00 00 00 07 58 22 0C 3A FE 28 71 BE DC 01 43 BD 84 71 6C 45 6C 80 8A 05
        """;

    public const string SpecRevisionManifest =
        """
        0C 56 0C 05 A9 D1 DF 9C 9B 2E 42 B2 59 81 7A F3 51 14 54 80 47 AF 30 54 71 6E 9B 40 98 06 70 7E 81 8D C1 02 34
        00 00 00 00 00 00 00 09 D0 24 0C 3A FE 28 71 BE DC 01 43 BD 84 71 6C 45 6C 80 8A 00 50 4C 14 B9 FA DE 84 A3 AA
        0D 4A A3 A8 52 0C 77 AC 70 73 80 13 38 0C DE AF 7C 55 4E 95 0E 65 7A D3 A3 FA 63 01 00 00 11 C8 22 0C 8E FC 0B
        2C 04 9B 61 4C AB 49 48 45 E6 03 EC A0 05
        """;

    public const string SpecStorageIndex =
        """
        0C 56 0C 8E 2E 2E 05 D1 C0 86 48 9C 51 29 D6 61 71 4F 67 80 0A 4E D0 67 25 4F E5 43 91 48 B7 28 D3 AB 89 77 01
        00 00 00 00 00 00 00 03 88 54 0C 99 FA 30 D7 2C 12 88 42 B7 22 0A 12 5C FD A7 E5 80 B8 50 CF AB 8E 91 64 BF 98
        06 70 7E 81 8D C1 02 3E 00 00 00 00 00 00 00 70 9A 0C B9 FA DE 84 A3 AA 0D 4A A3 A8 52 0C 77 AC 70 73 0C 65 46
        2A 6F C8 42 C7 46 BA B4 E2 8F DC E1 E3 2B 60 0C 8E FC 0B 2C 04 9B 61 4C AB 49 48 45 E6 03 EC A0 80 B8 50 CF AB
        8E 91 64 BF 98 06 70 7E 81 8D C1 02 40 00 00 00 00 00 00 00 68 76 0C 3A FE 28 71 BE DC 01 43 BD 84 71 6C 45 6C
        80 8A 0C 05 A9 D1 DF 9C 9B 2E 42 B2 59 81 7A F3 51 14 54 80 B8 50 CF AB 8E 91 64 BF 98 06 70 7E 81 8D C1 02 3F
        00 00 00 00 00 00 00 05
        """;

    /// <summary>
    /// The Put Changes request of [MS-FSSHTTPB] 4.3, without its Object Group, whose data the spec leaves out.
    /// </summary>
    public static byte[] SpecPutChangesRequest =>
    [
        ..Fsshttpb.Hex(SpecPutChangesRequestHeader),
        ..Fsshttpb.Hex(SpecStorageManifest),
        ..Fsshttpb.Hex(SpecCellManifest),
        ..Fsshttpb.Hex(SpecRevisionManifest),
        ..Fsshttpb.Hex(SpecStorageIndex),
        0x55,
        0x03,
        0x01
    ];

    public const string FsshttpQueryChangesRequest =
        "DAALAJzPKfM5lAabBgIAAO4CAACqAiAAfrgx50XdqkSrgAx1+9FTDnoCCAD0J0UfdwEWAgYAAwUAigICAADaAgYAAwAAygIIAAgAgAOEAEELAawCAFUDAQ==";

    public const string FsshttpQueryChangesRequestWithKnowledge =
        "DAALAJzPKfM5lAabBgIAAO4CAACqAiAAfrgx50XdqkSrgAx1+9FTDnoCCAD0J0UfdwEWAgYAAwUAigICAADaAgYAAwAAygIIAAgAgAOEACYCIAD2NXoyYQcURJaGUekAZnpNpAB4JE1n0+yUtD6/lXGrhF54W34Ac3gkspgsE2tLwUCVcauEXnhbfgBrURMBJgIgABMfCRCCyPtAmIZlM/k0wh1sAXAtDPkLQTdv0ZlEpsMnIy7cpxEJMgAAALUTASYCIAAO6XY6MoAMTbnd88ZQKUM+TAEgJgyymCwTa0vBQJVxq4ReeFt+awClEwFBCwGsAgBVAwE=";

    public static string[] FsshttpResponses =
    [
        "DAALAJ3PKfM5lAabFgMCAACsAgAMVgy19aRX2CXZQo8Tj3agbNI7gHJo8x4AR9pJiSDiCMP7d3ABAAAAAAAAAAOIVAz7/6ntX2CcT4gqqc6gT8RPgGskUtRR7xZFtwAl/eVQjrACAAAAAAAAAAUMVgz7/6ntX2CcT4gqqc6gT8RPgGskUtRR7xZFtwAl/eVQjrABAAAAAAAAAAVgIAAAAAAAAAAAAAAAAAAAAAAFVQ4CBgADBQD6AiQADLX1pFfYJdlCjxOPdqBs0jsAhABBBwGLAQ==",
        "DAALAJ3PKfM5lAabFgMCAACsAgAMVgxdgv0sL6qXQpGLEnn1TeJGgEXxEmphk/RDliumkvzDIGwBAAAAAAAAAAMFVQ4CBgADBQD6AiQADF2C/SwvqpdCkYsSefVN4kYAhAAmAiAA9jV6MmEHFESWhlHpAGZ6TaQAeCRNZ9PslLQ+v5Vxq4ReeFt+AHN4JLKYLBNrS8FAlXGrhF54W34Aa1ETASYCIAATHwkQgsj7QJiGZTP5NMIdbAFwLQz5C0E3b9GZRKbDJyMu3KcRCTIAAAC1EwEmAiAADul2OjKADE253fPGUClDPkwBICYMspgsE2tLwUCVcauEXnhbfmsApRMBQQcBiwE=",
        "DAALAJ3PKfM5lAabFgMCAACsAgAMVgx6KmdPAwAaQpf2Pb9Nlj2NgCminTMNs9JDutyBtjY1HeoBAAAAAAAAAAOIVAxZA/FHJtJ3R6iV1MCVAWthgNcfQoujYe9GnWDh5152y9MCAAAAAAAAAAUMVgxZA/FHJtJ3R6iV1MCVAWthgNcfQoujYe9GnWDh5152y9MBAAAAAAAAAAVgIAAAAAAAAAAAAAAAAAAAAAAFVQ4CBgADBQD6AiQADHoqZ08DABpCl/Y9v02WPY0AhABBBwGLAQ=="
    ];

    /// <summary>
    /// Word's Cell sub-requests: the editors table and content partitions, the file properties (with the cell roundtrip
    /// options), and the lone Query Access, with no user agent.
    /// </summary>
    public static string[] WordRequests =
    [
        "DgALAJzPKfM5lAabBgIAAO4CAACqAiAAjBCEGZNL606zIJGUMtblk1oEFgANbXN3b3JkB3dpbnoCCABm9qA5dwEWAgYAAwUAigIEAACs2gIGAAMAAMoCCAAIAIADhABBCwGsAgBVAwE=",
        "DgALAJzPKfM5lAabBgIAAO4CAACqAiAAjBCEGZNL606zIJGUMtblk1oEFgANbXN3b3JkB3dpbnoCCABm9qA5dwEWAgYAAwUAigIEAABQ2gIGAAMAAMoCCAAIAIADhABBCwGsAgBVAwE=",
        "DgALAJzPKfM5lAabBgIAAO4CAACqAiAAjBCEGZNL606zIJGUMtblk1oEFgANbXN3b3JkB3dpbnoCCABm9qA5dwFqBAIAARYCBgADBQCKAgQAABDaAgYAAwAAygIIAAgAgAOEAEELAawCAFUDAQ==",
        "DgALAJzPKfM5lAabBgIAABYCBgADAwALAQMB"
    ];

    /// <summary>
    /// Every vector, with the parser of the whole of it.
    /// </summary>
    public static IEnumerable<(string Name, byte[] Bytes, Func<byte[], object> Parse)> All()
    {
        yield return ("SpecQueryChangesRequest", Fsshttpb.Hex(SpecQueryChangesRequest), Fsshttpb.ParseRequest);
        yield return ("SpecQueryChangesSubResponse", Fsshttpb.Hex(SpecQueryChangesSubResponse), Fsshttpb.ParseSubResponse);
        yield return ("SpecPutChangesResponse", Fsshttpb.Hex(SpecPutChangesResponse), Fsshttpb.ParseResponse);
        yield return ("SpecPutChangesRequest", SpecPutChangesRequest, Fsshttpb.ParseRequest);
        yield return ("SpecStorageManifest", Fsshttpb.Hex(SpecStorageManifest), Fsshttpb.ParseElement);
        yield return ("SpecCellManifest", Fsshttpb.Hex(SpecCellManifest), Fsshttpb.ParseElement);
        yield return ("SpecRevisionManifest", Fsshttpb.Hex(SpecRevisionManifest), Fsshttpb.ParseElement);
        yield return ("SpecStorageIndex", Fsshttpb.Hex(SpecStorageIndex), Fsshttpb.ParseElement);
        yield return ("FsshttpQueryChangesRequest", Convert.FromBase64String(FsshttpQueryChangesRequest), Fsshttpb.ParseRequest);
        yield return ("FsshttpQueryChangesRequestWithKnowledge", Convert.FromBase64String(FsshttpQueryChangesRequestWithKnowledge), Fsshttpb.ParseRequest);
        for (var index = 0; index < FsshttpResponses.Length; index++)
        {
            yield return ($"FsshttpResponse{index}", Convert.FromBase64String(FsshttpResponses[index]), Fsshttpb.ParseResponse);
        }

        for (var index = 0; index < WordRequests.Length; index++)
        {
            yield return ($"WordRequest{index}", Convert.FromBase64String(WordRequests[index]), Fsshttpb.ParseRequest);
        }
    }
}

/// <summary>
/// Malformed FSSHTTPB: the reader throws <see cref="FsshttpbException"/> and nothing else, never hangs, and never
/// allocates beyond the size of its input. What it does read, the writer writes back in a form that reads the same.
/// </summary>
public class FsshttpbFuzzTests
{
    static readonly Guid guid = Fsshttpb.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073");

    static ExtendedGuid Id(uint value) => new(guid, value);

    /// <summary>
    /// Parses, and when that succeeds, checks that the model writes and reads back to itself. Returns whether the input
    /// was read: any exception but <see cref="FsshttpbException"/> fails, with the input in the message.
    /// </summary>
    static bool Check(string name, Func<byte[], object> parse, byte[] bytes)
    {
        object parsed;
        try
        {
            parsed = parse(bytes);
        }
        catch (FsshttpbException)
        {
            return false;
        }
        catch (Exception exception)
        {
            throw new($"{name}: {exception.GetType().Name} on {Fsshttpb.Hex(bytes)}", exception);
        }

        try
        {
            var written = Fsshttpb.Write(parsed);
            var again = parse(written);
            if (Fsshttpb.Describe(again) != Fsshttpb.Describe(parsed) ||
                Fsshttpb.Hex(Fsshttpb.Write(again)) != Fsshttpb.Hex(written))
            {
                throw new($"{name}: not written back the same");
            }
        }
        catch (Exception exception)
        {
            throw new($"{name}: the model of {Fsshttpb.Hex(bytes)} does not round trip", exception);
        }

        return true;
    }

    public static IEnumerable<string> Vectors() => FsshttpbVectors.All().Select(_ => _.Name);

    /// <summary>
    /// Every truncation of each vector, and every byte of it with each of its bits flipped, and all of them.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(Vectors))]
    public async Task Mutations(string name)
    {
        var (_, bytes, parse) = FsshttpbVectors.All().Single(_ => _.Name == name);
        var read = 0;
        var count = 0;
        for (var length = 0; length < bytes.Length; length++)
        {
            count++;
            if (Check(name, parse, bytes[..length]))
            {
                read++;
            }
        }

        byte[] masks = [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0xFF];
        for (var position = 0; position < bytes.Length; position++)
        {
            foreach (var mask in masks)
            {
                var mutated = bytes.ToArray();
                mutated[position] ^= mask;
                count++;
                if (Check(name, parse, mutated))
                {
                    read++;
                }
            }
        }

        // most mutations break the structure; the ones in data, GUIDs and values do not
        await Assert.That(read).IsLessThan(count);
        await Assert.That(Check(name, parse, bytes)).IsTrue();
    }

    /// <summary>
    /// Random bytes, alone and after the start of a request or a response, and spliced into the vectors, through every
    /// parser.
    /// </summary>
    [Test]
    public async Task Random()
    {
        var random = new Random(2026);
        Func<byte[], object>[] parsers =
        [
            Fsshttpb.ParseRequest,
            Fsshttpb.ParseResponse,
            Fsshttpb.ParseSubRequest,
            Fsshttpb.ParseSubResponse,
            Fsshttpb.ParseElement,
            Fsshttpb.ParseKnowledge,
            Fsshttpb.ParseError
        ];
        var vectors = FsshttpbVectors.All().Select(_ => _.Bytes).ToList();
        // bytes that are often stream object headers or the forms of the values
        byte[] common = [0x00, 0x01, 0x03, 0x04, 0x05, 0x0C, 0x14, 0x41, 0x56, 0x80, 0x84, 0x02, 0x06, 0x0E, 0x16, 0xAC, 0x55];
        var request = Fsshttpb.Hex("0C000B009CCF29F33994069B06020000");
        var response = Fsshttpb.Hex("0C000B009DCF29F33994069B1603020000");
        var read = 0;
        for (var iteration = 0; iteration < 20000; iteration++)
        {
            var tail = new byte[random.Next(0, 80)];
            for (var index = 0; index < tail.Length; index++)
            {
                if (random.Next(2) == 0)
                {
                    tail[index] = common[random.Next(common.Length)];
                }
                else
                {
                    tail[index] = (byte) random.Next(256);
                }
            }

            byte[] bytes;
            switch (iteration % 4)
            {
                case 0:
                    bytes = tail;
                    break;
                case 1:
                    bytes = [..request, ..tail];
                    break;
                case 2:
                    bytes = [..response, ..tail];
                    break;
                default:
                    var vector = vectors[random.Next(vectors.Count)];
                    var at = random.Next(vector.Length);
                    bytes = [..vector[..at], ..tail, ..vector[Math.Min(vector.Length, at + random.Next(8))..]];
                    break;
            }

            foreach (var parse in parsers)
            {
                if (Check($"random {iteration}", parse, bytes))
                {
                    read++;
                }
            }
        }

        await Assert.That(read).IsGreaterThan(0);
    }

    static long Allocated(Func<byte[], object> parse, byte[] bytes)
    {
        var before = GC.GetAllocatedBytesForCurrentThread();
        try
        {
            parse(bytes);
        }
        catch (FsshttpbException)
        {
        }

        return GC.GetAllocatedBytesForCurrentThread() - before;
    }

    /// <summary>
    /// Counts and lengths of up to 2^64 in a few bytes fail before allocating: the arrays of a Put Changes response, the
    /// Binary Item of a hash, the Large Length of a data element.
    /// </summary>
    [Test]
    public async Task HugeCounts()
    {
        var addedCount = Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.SubResponse, 3);
            _.WriteCompactUInt64(1);
            _.WriteCompactUInt64((ulong) RequestType.PutChanges);
            _.WriteByte(0);
            _.WriteStart(StreamObjectType.PutChangesResponse, 10);
            _.WriteExtendedGuid(ExtendedGuid.Null);
            _.WriteCompactUInt64(ulong.MaxValue);
        });
        var hashLength = Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.DataElement, 3);
            _.WriteExtendedGuid(ExtendedGuid.Null);
            _.WriteSerialNumber(SerialNumber.Null);
            _.WriteCompactUInt64((ulong) DataElementType.ObjectGroup);
            _.WriteStart(StreamObjectType.DataElementHash, 7);
            _.WriteCompactUInt64(1);
            _.WriteCompactUInt64(1UL << 40);
        });
        var largeLength = Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.DataElement, 3);
            _.WriteExtendedGuid(ExtendedGuid.Null);
            _.WriteSerialNumber(SerialNumber.Null);
            _.WriteCompactUInt64((ulong) DataElementType.ObjectDataBlob);
            _.WriteStart(StreamObjectType.ObjectDataBlob, ulong.MaxValue);
        });
        var supplementalLength = Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            ResponseError.Cell(CellErrorCode.Unknown).Write(ref _);
        });
        // the String Item count of the supplemental info, put where the end of the error was
        supplementalLength =
        [
            ..supplementalLength[..^2],
            ..Fsshttpb.Bytes((ref FsshttpbWriter _) =>
            {
                _.WriteStart(StreamObjectType.ErrorStringSupplementalInfo, 9);
                _.WriteCompactUInt64(1UL << 62);
            })
        ];

        await Assert.That(() => Fsshttpb.ParseSubResponse(addedCount)).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.ParseElement(hashLength)).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.ParseElement(largeLength)).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.ParseError(supplementalLength)).Throws<FsshttpbException>();
        foreach (var (parse, bytes) in new (Func<byte[], object>, byte[])[]
                 {
                     (Fsshttpb.ParseSubResponse, addedCount),
                     (Fsshttpb.ParseElement, hashLength),
                     (Fsshttpb.ParseElement, largeLength),
                     (Fsshttpb.ParseError, supplementalLength)
                 })
        {
            await Assert.That(Allocated(parse, bytes)).IsLessThan(64 * 1024);
        }
    }

    static byte[] UnknownKnowledge(int depth) =>
        Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.Knowledge, 0);
            _.WriteStart(StreamObjectType.SpecializedKnowledge, 16);
            _.WriteGuid(guid);
            for (var level = 0; level < depth; level++)
            {
                _.WriteStart(StreamObjectType.CellKnowledge, 0);
            }

            for (var level = 0; level < depth; level++)
            {
                _.WriteEnd(StreamObjectType.CellKnowledge);
            }

            _.WriteEnd(StreamObjectType.SpecializedKnowledge);
            _.WriteEnd(StreamObjectType.Knowledge);
        });

    /// <summary>
    /// Specialized Knowledge of an unknown kind is kept as its stream objects, which are skipped up to 32 levels deep.
    /// </summary>
    [Test]
    public async Task DeepNesting()
    {
        var shallow = UnknownKnowledge(33);
        var knowledge = (Knowledge) Fsshttpb.ParseKnowledge(shallow);
        var unknown = (UnknownKnowledge) knowledge.Items.Single();
        await Assert.That(unknown.Kind).IsEqualTo(guid);
        await Assert.That(unknown.Data.Length).IsEqualTo(33 * 3);
        await Assert.That(Fsshttpb.Write(knowledge)).IsEquivalentTo(shallow);

        var exception = Assert.Throws<FsshttpbException>(() => Fsshttpb.ParseKnowledge(UnknownKnowledge(34)));
        await Assert.That(exception.Message).StartsWith("Stream objects nested too deeply");
        await Assert.That(() => Fsshttpb.ParseKnowledge(UnknownKnowledge(100_000))).Throws<FsshttpbException>();
    }

    static ResponseError Chain(int length)
    {
        ResponseError? error = null;
        for (var index = 0; index < length; index++)
        {
            error = new(ResponseErrorType.Win32, (uint) index, null, error);
        }

        return error!;
    }

    /// <summary>
    /// Chained errors are read up to 9 deep.
    /// </summary>
    [Test]
    public async Task ErrorChain()
    {
        var chain = Chain(9);
        var bytes = Fsshttpb.Write(chain);
        await Assert.That(Fsshttpb.Describe(Fsshttpb.ParseError(bytes))).IsEqualTo(Fsshttpb.Describe(chain));
        var exception = Assert.Throws<FsshttpbException>(() => Fsshttpb.ParseError(Fsshttpb.Write(Chain(10))));
        await Assert.That(exception.Message).StartsWith("Response errors chained too deeply");
    }

    static byte[] Element(ulong type, Writing body) =>
        Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.DataElement, 3);
            _.WriteExtendedGuid(ExtendedGuid.Null);
            _.WriteSerialNumber(SerialNumber.Null);
            _.WriteCompactUInt64(type);
            body(ref _);
            _.WriteEnd(StreamObjectType.DataElement);
        });

    static byte[] ObjectGroup(Writing declarations, Writing data) =>
        Element((ulong) DataElementType.ObjectGroup, (ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.ObjectGroupDeclarations, 0);
            declarations(ref _);
            _.WriteEnd(StreamObjectType.ObjectGroupDeclarations);
            _.WriteStart(StreamObjectType.ObjectGroupData, 0);
            data(ref _);
            _.WriteEnd(StreamObjectType.ObjectGroupData);
        });

    static void Nothing(ref FsshttpbWriter writer)
    {
    }

    static void Blob(ref FsshttpbWriter writer)
    {
        writer.WriteStart(StreamObjectType.ObjectDataBlob, 1);
        writer.WriteByte(0);
    }

    static byte[] SubRequest(ulong type) =>
        Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.SubRequest, 3);
            _.WriteCompactUInt64(1);
            _.WriteCompactUInt64(type);
            _.WriteCompactUInt64(0);
            _.WriteEnd(StreamObjectType.SubRequest);
        });

    static byte[] SubResponse(ulong type) =>
        Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.SubResponse, 3);
            _.WriteCompactUInt64(1);
            _.WriteCompactUInt64(type);
            _.WriteByte(0);
            _.WriteEnd(StreamObjectType.SubResponse);
        });

    /// <summary>
    /// Structures that are well formed stream objects, but not what the spec allows where they are.
    /// </summary>
    [Test]
    public async Task Invalid()
    {
        var request = new FsshttpbRequest(null, [new QueryAccessRequest(1)]).ToArray();
        var response = new FsshttpbResponse([]).ToArray();
        var cases = new (string Name, Func<byte[], object> Parse, byte[] Bytes, string Message)[]
        {
            ("unknown data element type", Fsshttpb.ParseElement, Element(7, Nothing), "Unknown data element type 7"),
            ("object in a Storage Index", Fsshttpb.ParseElement, Element((ulong) DataElementType.StorageIndex, Blob), "in a Storage Index"),
            ("Object Group declaration", Fsshttpb.ParseElement, ObjectGroup(Blob, Nothing), "in Object Group declarations"),
            ("Object Group data", Fsshttpb.ParseElement, ObjectGroup(Nothing, Blob), "in Object Group data"),
            ("unknown sub-request type", Fsshttpb.ParseSubRequest, SubRequest(3), "Unknown request type 3"),
            ("unknown sub-response type", Fsshttpb.ParseSubResponse, SubResponse(3), "Unknown request type 3"),
            ("request signature", Fsshttpb.ParseRequest, [..request[..4], 0, ..request[5..]], "Not an FSSHTTPB request signature"),
            ("response signature", Fsshttpb.ParseResponse, request, "Not an FSSHTTPB response signature"),
            ("data after the request", Fsshttpb.ParseRequest, [..request, 0], "1 unexpected bytes at the end of the request"),
            ("data after the response", Fsshttpb.ParseResponse, [..response, 0], "1 unexpected bytes at the end of the response"),
            ("Query Changes flags", Fsshttpb.ParseSubRequest, QueryChangesFlags(3), "Query Changes flags of 3 bytes"),
            ("Query Changes flags", Fsshttpb.ParseSubRequest, QueryChangesFlags(0), "Query Changes flags of 0 bytes"),
            ("error type", Fsshttpb.ParseError, UnknownError(), "Unknown error type"),
            ("object in Cell Knowledge", Fsshttpb.ParseKnowledge, CellKnowledgeWithBlob(), "in Cell Knowledge")
        };
        foreach (var (name, parse, bytes, message) in cases)
        {
            var exception = Assert.Throws<FsshttpbException>(() => parse(bytes));
            await Assert.That(exception.Message).Contains(message);
        }
    }

    static byte[] QueryChangesFlags(int length) =>
        Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.SubRequest, 3);
            _.WriteCompactUInt64(1);
            _.WriteCompactUInt64((ulong) RequestType.QueryChanges);
            _.WriteCompactUInt64(0);
            _.WriteStart(StreamObjectType.QueryChangesRequest, length);
            for (var index = 0; index < length; index++)
            {
                _.WriteByte(0);
            }

            _.WriteEnd(StreamObjectType.SubRequest);
        });

    static byte[] UnknownError()
    {
        var bytes = Fsshttpb.Write(ResponseError.Success);
        // the first byte of the GUID of the error type, after its 32-bit start
        bytes[4] ^= 0xFF;
        return bytes;
    }

    static byte[] CellKnowledgeWithBlob() =>
        Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.Knowledge, 0);
            _.WriteStart(StreamObjectType.SpecializedKnowledge, 16);
            _.WriteGuid(SpecializedKnowledge.CellKnowledgeKind);
            _.WriteStart(StreamObjectType.CellKnowledge, 0);
            Blob(ref _);
            _.WriteEnd(StreamObjectType.CellKnowledge);
            _.WriteEnd(StreamObjectType.SpecializedKnowledge);
            _.WriteEnd(StreamObjectType.Knowledge);
        });

    /// <summary>
    /// A model the writer cannot write: an Extended GUID or Serial Number with a value and the empty GUID.
    /// </summary>
    [Test]
    public async Task Write_Invalid()
    {
        await Assert.That(() => Fsshttpb.Write(new CellManifest(new(Guid.Empty, 1), SerialNumber.Null, Id(1)))).Throws<ArgumentException>();
        await Assert.That(() => Fsshttpb.Write(new CellManifest(Id(1), new(Guid.Empty, 1), Id(1)))).Throws<ArgumentException>();
    }
}

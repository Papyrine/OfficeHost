using System.Numerics;

/// <summary>
/// The basic data types of [MS-FSSHTTPB] 2.2.1: compact integers, Extended GUIDs, Serial Numbers, the items and arrays,
/// and the stream object headers.
/// </summary>
public class FsshttpbPrimitiveTests
{
    static readonly Guid guid = Fsshttpb.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073");
    const string guidHex = "B9FADE84A3AA0D4AA3A8520C77AC7073";

    static string Compact(ulong value) => Fsshttpb.Hex(Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteCompactUInt64(value)));

    static ulong ReadCompact(string hex) => Fsshttpb.Read(Fsshttpb.Hex(hex), (ref FsshttpbReader _) => _.ReadCompactUInt64());

    /// <summary>
    /// 2.2.1.1: the encodings of the spec examples (4.1 to 4.4) and of each size at its limits.
    /// </summary>
    [Test]
    [Arguments(0ul, "00")]
    [Arguments(1ul, "03")]
    [Arguments(2ul, "05")]
    [Arguments(5ul, "0B")]
    [Arguments(111ul, "DF")]
    [Arguments(116ul, "E9")]
    [Arguments(0x7Ful, "FF")]
    [Arguments(0x80ul, "0202")]
    [Arguments(0x3FFFul, "FEFF")]
    [Arguments(0x4000ul, "040002")]
    [Arguments(73507ul, "1CF908")]
    [Arguments(0x1F_FFFFul, "FCFFFF")]
    [Arguments(0x20_0000ul, "08000002")]
    [Arguments(3670016ul, "08008003")]
    [Arguments(0xFFF_FFFFul, "F8FFFFFF")]
    [Arguments(0x1000_0000ul, "1000000002")]
    [Arguments(0x7_FFFF_FFFFul, "F0FFFFFFFF")]
    [Arguments(0x8_0000_0000ul, "200000000002")]
    [Arguments(0x3FF_FFFF_FFFFul, "E0FFFFFFFFFF")]
    [Arguments(0x400_0000_0000ul, "40000000000002")]
    [Arguments(0x1_FFFF_FFFF_FFFFul, "C0FFFFFFFFFFFF")]
    [Arguments(0x2_0000_0000_0000ul, "800000000000000200")]
    [Arguments(ulong.MaxValue, "80FFFFFFFFFFFFFFFF")]
    public async Task CompactUInt64_Encoding(ulong value, string hex)
    {
        await Assert.That(Compact(value)).IsEqualTo(hex);
        await Assert.That(ReadCompact(hex)).IsEqualTo(value);
        await Assert.That(CompactUInt64.Size(value)).IsEqualTo(hex.Length / 2);
    }

    /// <summary>
    /// 2.2.1.1: random values of every bit length, and the values around each limit, read back as written, in the
    /// shortest encoding.
    /// </summary>
    [Test]
    public async Task CompactUInt64_Property()
    {
        var random = new Random(2026);
        var values = new List<ulong>();
        for (var bits = 0; bits <= 64; bits++)
        {
            var limit = ulong.MaxValue;
            if (bits < 64)
            {
                limit = (1UL << bits) - 1;
            }

            values.AddRange([limit, limit - 1, limit + 1]);
            for (var index = 0; index < 50; index++)
            {
                values.Add((ulong) random.NextInt64() & limit);
            }
        }

        foreach (var value in values)
        {
            var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteCompactUInt64(value));
            await Assert.That(bytes.Length).IsEqualTo(CompactUInt64.Size(value));
            await Assert.That(bytes.Length).IsEqualTo(ExpectedCompactSize(value));
            await Assert.That(Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadCompactUInt64())).IsEqualTo(value);
        }
    }

    static int ExpectedCompactSize(ulong value)
    {
        var bits = 64 - BitOperations.LeadingZeroCount(value);
        if (bits > 49)
        {
            return 9;
        }

        return Math.Max(1, (bits + 6) / 7);
    }

    /// <summary>
    /// 2.2.1.1: a value written longer than needed is read. There is no 8-byte form: a first byte of 128 has 8 bytes of
    /// value after it.
    /// </summary>
    [Test]
    [Arguments("0600")]
    [Arguments("0C0000")]
    [Arguments("800100000000000000")]
    public async Task CompactUInt64_LongerThanNeeded(string hex) =>
        await Assert.That(ReadCompact(hex)).IsEqualTo(1ul);

    [Test]
    [Arguments("")]
    [Arguments("02")]
    [Arguments("0C00")]
    [Arguments("80FFFFFFFFFFFFFF")]
    [Arguments("C0FFFFFFFFFF")]
    public async Task CompactUInt64_Truncated(string hex) =>
        await Assert.That(() => ReadCompact(hex)).Throws<FsshttpbException>();

    static string ExtendedGuidHex(ExtendedGuid value) =>
        Fsshttpb.Hex(Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteExtendedGuid(value)));

    static ExtendedGuid ReadExtendedGuid(string hex) =>
        Fsshttpb.Read(Fsshttpb.Hex(hex), (ref FsshttpbReader _) => _.ReadExtendedGuid());

    /// <summary>
    /// 2.2.1.7: the forms of the spec examples (the 5-bit 0C and 14, the 10-bit 600C, the 32-bit 80 13380CDE of 4.3.5),
    /// and each form at its limits.
    /// </summary>
    [Test]
    [Arguments(0u, "04")]
    [Arguments(1u, "0C")]
    [Arguments(2u, "14")]
    [Arguments(0x1Fu, "FC")]
    [Arguments(0x20u, "2008")]
    [Arguments(0x31u, "600C")]
    [Arguments(0x3FFu, "E0FF")]
    [Arguments(0x400u, "400002")]
    [Arguments(0x1FFFFu, "C0FFFF")]
    [Arguments(0x20000u, "8000000200")]
    [Arguments(0xDE0C3813u, "8013380CDE")]
    [Arguments(uint.MaxValue, "80FFFFFFFF")]
    public async Task ExtendedGuid_Encoding(uint value, string prefix)
    {
        var extended = new ExtendedGuid(guid, value);
        var hex = prefix + guidHex;
        await Assert.That(ExtendedGuidHex(extended)).IsEqualTo(hex);
        await Assert.That(ReadExtendedGuid(hex)).IsEqualTo(extended);
        await Assert.That(extended.Size).IsEqualTo(hex.Length / 2);
    }

    [Test]
    public async Task ExtendedGuid_Null()
    {
        await Assert.That(ExtendedGuidHex(ExtendedGuid.Null)).IsEqualTo("00");
        await Assert.That(ReadExtendedGuid("00").IsNull).IsTrue();
        await Assert.That(ExtendedGuid.Null.Size).IsEqualTo(1);
        await Assert.That(new ExtendedGuid(guid, 0).IsNull).IsFalse();
        await Assert.That(ExtendedGuid.Null.ToString()).IsEqualTo("{00000000-0000-0000-0000-000000000000},0");
    }

    /// <summary>
    /// 2.2.1.7: random values read back as written, in the shortest form.
    /// </summary>
    [Test]
    public async Task ExtendedGuid_Property()
    {
        var random = new Random(2026);
        uint[] limits = [0x1F, 0x3FF, 0x1FFFF, uint.MaxValue];
        foreach (var limit in limits)
        {
            for (var index = 0; index < 50; index++)
            {
                var value = new ExtendedGuid(Guid.NewGuid(), (uint) random.NextInt64(0, (long) limit + 1));
                var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteExtendedGuid(value));
                await Assert.That(bytes.Length).IsEqualTo(value.Size);
                await Assert.That(Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadExtendedGuid())).IsEqualTo(value);
            }
        }
    }

    /// <summary>
    /// 2.2.1.7: a first byte that is none of the forms, the empty GUID in a form that is not null, and truncation.
    /// </summary>
    [Test]
    [Arguments("01")]
    [Arguments("02")]
    [Arguments("08")]
    [Arguments("81")]
    [Arguments("C0")]
    [Arguments("0C00000000000000000000000000000000")]
    [Arguments("0CB9FADE84A3AA0D4AA3A8520C77AC70")]
    [Arguments("80FFFF")]
    [Arguments("")]
    public async Task ExtendedGuid_Invalid(string hex)
    {
        var exception = Assert.Throws<FsshttpbException>(() => ReadExtendedGuid(hex));
        await Assert.That(exception.Message).Contains("(at byte");
    }

    [Test]
    public async Task ExtendedGuid_EmptyGuidWithValue()
    {
        var exception = Assert.Throws<ArgumentException>(() => ExtendedGuidHex(new(Guid.Empty, 1)));
        await Assert.That(exception.ParamName).IsEqualTo("value");
    }

    [Test]
    public async Task SerialNumber_Encoding()
    {
        var value = new SerialNumber(guid, 0x32);
        var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteSerialNumber(value));
        await Assert.That(Fsshttpb.Hex(bytes)).IsEqualTo("80" + guidHex + "3200000000000000");
        await Assert.That(value.Size).IsEqualTo(25);
        await Assert.That(Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadSerialNumber())).IsEqualTo(value);
        await Assert.That(value.ToString()).IsEqualTo("{84defab9-aaa3-4a0d-a3a8-520c77ac7073},50");

        var empty = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteSerialNumber(SerialNumber.Null));
        await Assert.That(Fsshttpb.Hex(empty)).IsEqualTo("00");
        await Assert.That(SerialNumber.Null.Size).IsEqualTo(1);
        await Assert.That(Fsshttpb.Read(empty, (ref FsshttpbReader _) => _.ReadSerialNumber()).IsNull).IsTrue();
    }

    [Test]
    [Arguments("01")]
    [Arguments("40")]
    [Arguments("80000000000000000000000000000000000100000000000000")]
    [Arguments("80" + guidHex + "01")]
    public async Task SerialNumber_Invalid(string hex) =>
        await Assert.That(() => Fsshttpb.Read(Fsshttpb.Hex(hex), (ref FsshttpbReader _) => _.ReadSerialNumber())).Throws<FsshttpbException>();

    [Test]
    public async Task SerialNumber_EmptyGuidWithValue() =>
        await Assert.That(() => Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteSerialNumber(new(Guid.Empty, 1)))).Throws<ArgumentException>();

    /// <summary>
    /// 2.2.1.10, 2.2.1.11: a Cell ID is two Extended GUIDs, and the null Cell ID two null ones.
    /// </summary>
    [Test]
    public async Task CellId_Encoding()
    {
        var cell = new CellId(new(guid, 1), new(guid, 2));
        CellId[] cells = [cell, CellId.Null];
        var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteCellIdArray(cells));
        await Assert.That(Fsshttpb.Hex(bytes)).IsEqualTo("05" + "0C" + guidHex + "14" + guidHex + "0000");
        await Assert.That(FsshttpbSize.CellIdArray(cells)).IsEqualTo(bytes.Length);
        await Assert.That(CellId.Null.Size).IsEqualTo(2);
        var read = Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadCellIdArray());
        await Assert.That(read).IsEquivalentTo(cells);
    }

    /// <summary>
    /// 2.2.1.2: a File Chunk Reference is two compact integers.
    /// </summary>
    [Test]
    public async Task FileChunkReference_Encoding()
    {
        var chunk = new FileChunkReference(0x80, 3670016);
        var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteFileChunkReference(chunk));
        await Assert.That(Fsshttpb.Hex(bytes)).IsEqualTo("0202" + "08008003");
        await Assert.That(chunk.Size).IsEqualTo(6);
        await Assert.That(Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadFileChunkReference())).IsEqualTo(chunk);
    }

    /// <summary>
    /// 2.2.1.8: an Extended GUID Array is a count, then the Extended GUIDs.
    /// </summary>
    [Test]
    public async Task ExtendedGuidArray_Encoding()
    {
        ExtendedGuid[] values = [new(guid, 1), ExtendedGuid.Null, new(guid, 0x31)];
        var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteExtendedGuidArray(values));
        await Assert.That(Fsshttpb.Hex(bytes)).IsEqualTo("07" + "0C" + guidHex + "00" + "600C" + guidHex);
        await Assert.That(FsshttpbSize.ExtendedGuidArray(values)).IsEqualTo(bytes.Length);
        await Assert.That(Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadExtendedGuidArray())).IsEquivalentTo(values);
        var empty = Fsshttpb.Read(Fsshttpb.Hex("00"), (ref FsshttpbReader _) => _.ReadExtendedGuidArray());
        await Assert.That(empty.Length).IsEqualTo(0);
    }

    /// <summary>
    /// 2.2.1.3, 2.2.1.4, 2.2.1.14: a Binary Item is a count of bytes then the bytes, a String Item a count of UTF-16
    /// characters then the characters, with no terminator.
    /// </summary>
    [Test]
    public async Task Items_Encoding()
    {
        var binary = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteBinaryItem([1, 2, 3]));
        await Assert.That(Fsshttpb.Hex(binary)).IsEqualTo("07010203");
        await Assert.That(FsshttpbSize.BinaryItem(3)).IsEqualTo(4);
        var read = Fsshttpb.Read(binary, (ref FsshttpbReader _) => _.ReadBinaryItem());
        await Assert.That(Fsshttpb.Hex(read.Span)).IsEqualTo("010203");

        var text = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteStringItem("Wé"));
        await Assert.That(Fsshttpb.Hex(text)).IsEqualTo("055700E900");
        await Assert.That(FsshttpbSize.StringItem("Wé")).IsEqualTo(5);
        await Assert.That(Fsshttpb.Read(text, (ref FsshttpbReader _) => _.ReadStringItem())).IsEqualTo("Wé");

        string[] strings = ["a", "", "bc"];
        var array = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteStringItemArray(strings));
        await Assert.That(Fsshttpb.Hex(array)).IsEqualTo("07" + "036100" + "00" + "0562006300");
        await Assert.That(FsshttpbSize.StringItemArray(strings)).IsEqualTo(array.Length);
        await Assert.That(Fsshttpb.Read(array, (ref FsshttpbReader _) => _.ReadStringItemArray())).IsEquivalentTo(strings);
    }

    /// <summary>
    /// A count larger than the bytes left fails before anything is allocated.
    /// </summary>
    [Test]
    [Arguments("80FFFFFFFFFFFFFF7F")]
    [Arguments("8000000000000100000000")]
    [Arguments("0B")]
    public async Task Count_LargerThanData(string hex)
    {
        var bytes = Fsshttpb.Hex(hex);
        await Assert.That(() => Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadExtendedGuidArray())).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadCellIdArray())).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadBinaryItem())).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadStringItem())).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadStringItemArray())).Throws<FsshttpbException>();
    }

    [Test]
    public async Task FixedWidth_RoundTrip()
    {
        var value = Guid.NewGuid();
        var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteByte(0xAB);
            _.WriteUInt16(0x1234);
            _.WriteUInt32(0x12345678);
            _.WriteUInt64(0x1122334455667788);
            _.WriteGuid(value);
            _.WriteBytes([9, 8]);
        });
        await Assert.That(Fsshttpb.Hex(bytes.AsSpan(0, 15))).IsEqualTo("AB3412785634128877665544332211");
        var read = ReadFixedWidth(bytes);
        await Assert.That(read.ToString()).IsEqualTo((0xAB, 0x1234, 0x12345678, 0x1122334455667788, value, 2, "0908", true, false, false).ToString());
    }

    static (byte, ushort, uint, ulong, Guid, int, string, bool, bool, bool) ReadFixedWidth(byte[] bytes)
    {
        var reader = new FsshttpbReader(bytes);
        var oneByte = reader.ReadByte();
        var twoBytes = reader.ReadUInt16();
        var fourBytes = reader.ReadUInt32();
        var eightBytes = reader.ReadUInt64();
        var guid = reader.ReadGuid();
        var remaining = reader.Remaining;
        var rest = Fsshttpb.Hex(reader.ReadRest().Span);
        return (oneByte, twoBytes, fourBytes, eightBytes, guid, remaining, rest, reader.AtEnd, reader.PeekStart(StreamObjectType.Knowledge), reader.PeekEnd());
    }

    static StreamObjectHeader ReadHeader(byte[] bytes)
    {
        var reader = new FsshttpbReader(bytes);
        return reader.ReadHeader();
    }

    static StreamObjectHeader ReadHeader(string hex) =>
        Fsshttpb.Read(Fsshttpb.Hex(hex), (ref FsshttpbReader _) => _.ReadHeader());

    /// <summary>
    /// 2.2.1.5.1, 2.2.1.5.2: the starts of the spec examples, read and written back.
    /// </summary>
    [Test]
    [Arguments("06020000", (ushort) StreamObjectType.Request, 0ul)]
    [Arguments("EE020000", (ushort) StreamObjectType.UserAgent, 0ul)]
    [Arguments("7A020800", (ushort) StreamObjectType.UserAgentVersion, 4ul)]
    [Arguments("0E020600", (ushort) StreamObjectType.SubResponse, 3ul)]
    [Arguments("FA022400", (ushort) StreamObjectType.QueryChangesResponse, 18ul)]
    [Arguments("0C56", (ushort) StreamObjectType.DataElement, 43ul)]
    [Arguments("B200EC03", (ushort) StreamObjectType.ObjectGroupObjectData, 502ul)]
    [Arguments("D2022600", (ushort) StreamObjectType.PutChangesRequest, 19ul)]
    [Arguments("AC02", (ushort) StreamObjectType.DataElementPackage, 1ul)]
    [Arguments("8400", (ushort) StreamObjectType.Knowledge, 0ul)]
    public async Task Header_Start(string hex, ushort code, ulong length)
    {
        var type = (StreamObjectType) code;
        await Assert.That(ReadHeader(hex)).IsEqualTo(new StreamObjectHeader(type, type.IsCompound(), false, length));
        var written = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteStart(type, length));
        await Assert.That(Fsshttpb.Hex(written)).IsEqualTo(hex);
    }

    /// <summary>
    /// 2.2.1.5.3, 2.2.1.5.4: the ends of the spec examples, 8-bit below 0x40 and 16-bit from 0x40.
    /// </summary>
    [Test]
    [Arguments("05", (ushort) StreamObjectType.DataElement)]
    [Arguments("41", (ushort) StreamObjectType.Knowledge)]
    [Arguments("51", (ushort) StreamObjectType.CellKnowledge)]
    [Arguments("55", (ushort) StreamObjectType.DataElementPackage)]
    [Arguments("75", (ushort) StreamObjectType.ObjectGroupDeclarations)]
    [Arguments("79", (ushort) StreamObjectType.ObjectGroupData)]
    [Arguments("A5", (ushort) StreamObjectType.WaterlineKnowledge)]
    [Arguments("B5", (ushort) StreamObjectType.ContentTagKnowledge)]
    [Arguments("0301", (ushort) StreamObjectType.Request)]
    [Arguments("0701", (ushort) StreamObjectType.SubResponse)]
    [Arguments("0B01", (ushort) StreamObjectType.SubRequest)]
    [Arguments("1301", (ushort) StreamObjectType.SpecializedKnowledge)]
    [Arguments("8B01", (ushort) StreamObjectType.Response)]
    [Arguments("E701", (ushort) StreamObjectType.ObjectGroupMetadataDeclarations)]
    public async Task Header_End(string hex, ushort code)
    {
        var type = (StreamObjectType) code;
        await Assert.That(ReadHeader(hex)).IsEqualTo(new StreamObjectHeader(type, true, true, 0));
        var written = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteEnd(type));
        await Assert.That(Fsshttpb.Hex(written)).IsEqualTo(hex);
    }

    /// <summary>
    /// The writer uses a 16-bit start for a type below 0x40 and a length up to 127, a 32-bit start otherwise.
    /// </summary>
    [Test]
    [Arguments((ushort) StreamObjectType.QueryChangesVersioning, 127ul, "80FF")]
    [Arguments((ushort) StreamObjectType.QueryChangesVersioning, 128ul, "82010001")]
    [Arguments((ushort) StreamObjectType.DataElement, 0ul, "0C00")]
    [Arguments((ushort) StreamObjectType.Request, 0ul, "06020000")]
    [Arguments((ushort) StreamObjectType.ObjectDataBlob, 32766ul, "1200FCFF")]
    public async Task Header_Width(ushort code, ulong length, string hex)
    {
        var type = (StreamObjectType) code;
        var written = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteStart(type, length));
        await Assert.That(Fsshttpb.Hex(written)).IsEqualTo(hex);
        await Assert.That(ReadHeader(written)).IsEqualTo(new StreamObjectHeader(type, type.IsCompound(), false, length));
    }

    /// <summary>
    /// 2.2.1.5.2: a length from 32767 is a Large Length, a compact integer after the length field of 32767.
    /// </summary>
    [Test]
    [Arguments(32767ul, "1200FEFF" + "FCFF03")]
    [Arguments(40000ul, "1200FEFF" + "04E204")]
    [Arguments(1ul << 40, "1200FEFF" + "200000000040")]
    public async Task Header_LargeLength(ulong length, string hex)
    {
        var written = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteStart(StreamObjectType.ObjectDataBlob, length));
        await Assert.That(Fsshttpb.Hex(written)).IsEqualTo(hex);
        await Assert.That(ReadHeader(hex)).IsEqualTo(new StreamObjectHeader(StreamObjectType.ObjectDataBlob, false, false, length));
    }

    /// <summary>
    /// A start whose length is beyond the data fails, whatever the length, without allocating it.
    /// </summary>
    [Test]
    [Arguments(200ul)]
    [Arguments(32767ul)]
    [Arguments(ulong.MaxValue)]
    public async Task Header_LengthBeyondData(ulong length)
    {
        var bytes = Fsshttpb.Bytes((ref FsshttpbWriter _) =>
        {
            _.WriteStart(StreamObjectType.ObjectDataBlob, length);
            _.WriteBytes([1, 2, 3]);
        });
        var exception = Assert.Throws<FsshttpbException>(() => Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadStart(StreamObjectType.ObjectDataBlob).Remaining));
        await Assert.That(exception.Message).Contains("does not fit");
        await Assert.That(() => Fsshttpb.Read(bytes, (ref FsshttpbReader _) => _.ReadSingle())).Throws<FsshttpbException>();
    }

    static int Start(string hex, StreamObjectType type) =>
        Fsshttpb.Read(Fsshttpb.Hex(hex), (ref FsshttpbReader _) => _.ReadStart(type).Remaining);

    static int End(string hex, StreamObjectType type) =>
        Fsshttpb.Read(Fsshttpb.Hex(hex), (ref FsshttpbReader _) =>
        {
            _.ReadEnd(type);
            return 0;
        });

    /// <summary>
    /// A start of another type, an end where a start is expected, a compound bit that does not match the type, and an end
    /// of another type, all fail.
    /// </summary>
    [Test]
    public async Task Header_Mismatch()
    {
        await Assert.That(Start("8400", StreamObjectType.Knowledge)).IsEqualTo(0);
        await Assert.That(() => Start("8400", StreamObjectType.CellKnowledge)).Throws<FsshttpbException>();
        await Assert.That(() => Start("41", StreamObjectType.Knowledge)).Throws<FsshttpbException>();
        // Knowledge without its compound bit
        await Assert.That(() => Start("8000", StreamObjectType.Knowledge)).Throws<FsshttpbException>();
        // a single type with the compound bit
        await Assert.That(() => Start("1400", StreamObjectType.ObjectDataBlob)).Throws<FsshttpbException>();
        await Assert.That(End("41", StreamObjectType.Knowledge)).IsEqualTo(0);
        await Assert.That(() => End("41", StreamObjectType.CellKnowledge)).Throws<FsshttpbException>();
        await Assert.That(() => End("8400", StreamObjectType.Knowledge)).Throws<FsshttpbException>();
        // a single object is expected: neither a compound start, nor a compound type without its bit, nor an end
        await Assert.That(() => Fsshttpb.Read(Fsshttpb.Hex("8400"), (ref FsshttpbReader _) => _.ReadSingle())).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.Read(Fsshttpb.Hex("8000"), (ref FsshttpbReader _) => _.ReadSingle())).Throws<FsshttpbException>();
        await Assert.That(() => Fsshttpb.Read(Fsshttpb.Hex("41"), (ref FsshttpbReader _) => _.ReadSingle())).Throws<FsshttpbException>();
    }

    [Test]
    public async Task Header_ToString()
    {
        await Assert.That(new StreamObjectHeader(StreamObjectType.Knowledge, true, false, 0).ToString()).IsEqualTo("the start of Knowledge (0 bytes)");
        await Assert.That(new StreamObjectHeader(StreamObjectType.Knowledge, true, true, 0).ToString()).IsEqualTo("the end of Knowledge");
    }

    /// <summary>
    /// A single object, read as raw data and written back, and the peeks, which consume nothing.
    /// </summary>
    [Test]
    public async Task Single_RoundTrip()
    {
        var bytes = Fsshttpb.Hex("B200EC03" + new string('A', 1004));
        var (peeks, single, atEnd) = ReadSingle(bytes);
        await Assert.That(peeks).IsEqualTo((true, false, false, 0));
        await Assert.That(atEnd).IsTrue();
        await Assert.That(single.Type).IsEqualTo(StreamObjectType.ObjectGroupObjectData);
        await Assert.That(single.Data.Length).IsEqualTo(502);
        var written = Fsshttpb.Bytes((ref FsshttpbWriter _) => _.WriteSingle(single));
        await Assert.That(written).IsEquivalentTo(bytes);
    }

    static ((bool, bool, bool, int) Peeks, RawStreamObject Single, bool AtEnd) ReadSingle(byte[] bytes)
    {
        var reader = new FsshttpbReader(bytes);
        var peeks = (reader.PeekStart(StreamObjectType.ObjectGroupObjectData), reader.PeekStart(StreamObjectType.ObjectDataBlob), reader.PeekEnd(), reader.Offset);
        var single = reader.ReadSingle();
        return (peeks, single, reader.AtEnd);
    }

    /// <summary>
    /// Errors give the offset in the whole input, also from the reader of a part.
    /// </summary>
    [Test]
    public async Task Error_Offset()
    {
        var bytes = Fsshttpb.Hex("0C56" + "0C" + guidHex + new string('0', 82));
        var exception = Assert.Throws<FsshttpbException>(() => ReadPart(bytes));
        await Assert.That(exception.Offset).IsEqualTo(2 + 17 + 8);
        await Assert.That(exception.Message).IsEqualTo("18 unexpected bytes at the end of the test (at byte 27)");
    }

    static void ReadPart(byte[] bytes)
    {
        var reader = new FsshttpbReader(bytes);
        var fields = reader.ReadStart(StreamObjectType.DataElement);
        fields.ReadExtendedGuid();
        fields.ReadUInt64();
        fields.EnsureEnd("the test");
    }
}

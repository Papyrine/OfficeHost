/// <summary>
/// A Data Element Package ([MS-FSSHTTPB] 2.2.1.12): the data elements of a request or a response.
/// </summary>
record DataElementPackage(IReadOnlyList<DataElement> Elements)
{
    public static DataElementPackage Read(ref FssHttpBReader reader)
    {
        var fields = reader.ReadStart(StreamObjectType.DataElementPackage);
        // the Reserved byte, which is ignored
        fields.ReadByte();
        fields.EnsureEnd("the Data Element Package start");
        var elements = new List<DataElement>();
        while (reader.PeekStart(StreamObjectType.DataElement))
        {
            elements.Add(DataElement.Read(ref reader));
        }

        reader.ReadEnd(StreamObjectType.DataElementPackage);
        return new(elements);
    }

    public void Write(ref FssHttpBWriter writer)
    {
        writer.WriteStart(StreamObjectType.DataElementPackage, 1);
        writer.WriteByte(0);
        foreach (var element in Elements)
        {
            element.Write(ref writer);
        }

        writer.WriteEnd(StreamObjectType.DataElementPackage);
    }
}
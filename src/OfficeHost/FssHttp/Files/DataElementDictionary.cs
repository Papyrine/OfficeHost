static class DataElementDictionary
{
    public static Dictionary<ExtendedGuid, DataElement> With(this Dictionary<ExtendedGuid, DataElement> elements, IEnumerable<DataElement> added)
    {
        foreach (var element in added)
        {
            elements[element.Id] = element;
        }

        return elements;
    }
}
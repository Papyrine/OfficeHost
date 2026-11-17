/// <summary>
/// The XML names of the SOAP of Office ([MS-FSSHTTP] 2.2): the namespaces, and the names of the elements. A name made
/// from a namespace and a string is looked up in the namespace's table each time, so each is made once here.
/// </summary>
static class Soap
{
    public static XNamespace Namespace { get; } = "http://schemas.microsoft.com/sharepoint/soap/";
    public static XNamespace EnvelopeNamespace { get; } = "http://schemas.xmlsoap.org/soap/envelope/";
    public static XNamespace XopNamespace { get; } = "http://www.w3.org/2004/08/xop/include";

    public static XName Envelope { get; } = EnvelopeNamespace + "Envelope";
    public static XName Body { get; } = EnvelopeNamespace + "Body";

    // the reference of an MTOM part ([XOP10])
    public static XName Include { get; } = XopNamespace + "Include";

    public static XName Request { get; } = Namespace + "Request";
    public static XName SubRequest { get; } = Namespace + "SubRequest";
    public static XName SubRequestData { get; } = Namespace + "SubRequestData";
    public static XName Response { get; } = Namespace + "Response";
    public static XName SubResponse { get; } = Namespace + "SubResponse";
    public static XName SubResponseData { get; } = Namespace + "SubResponseData";
    public static XName ResponseVersion { get; } = Namespace + "ResponseVersion";
    public static XName ResponseCollection { get; } = Namespace + "ResponseCollection";

    // the properties of a document and its folder (Properties and GetDocMetaInfo sub-requests)
    public static XName Property { get; } = Namespace + "Property";
    public static XName PropertyValues { get; } = Namespace + "PropertyValues";
    public static XName DocProps { get; } = Namespace + "DocProps";
    public static XName FolderProps { get; } = Namespace + "FolderProps";

    // the answer of sharedaccess.asmx
    public static XName IsOnlyClientResponse { get; } = Namespace + "IsOnlyClientResponse";
    public static XName IsOnlyClientResult { get; } = Namespace + "IsOnlyClientResult";
}

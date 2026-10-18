/// <summary>
/// A local storage where each Reva token is a distinct user, as the CS3 storage maps them.
/// The local storage itself maps every user to root.
/// </summary>
sealed class MultiUserStorage(IStorage inner) :
    ForwardingStorage(inner)
{
    public override (string UserId, string WopiUser) GetUserIdFromCreds(string token, string wopiUser) =>
        (token, $"{wopiUser}!{token}");
}

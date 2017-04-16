// // --------------------------------------------------------------------------------------------------------------------
// // <copyright file="noor.alam.shuvo@gmail.com" company="">
// //   Copyright @ 2017
// // </copyright>
// <summary>
// // </summary>
// // --------------------------------------------------------------------------------------------------------------------

namespace AutoHourLogger
{
    using System;
    using System.IO;
    using System.Threading;

    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Util.Store;

    public static class UserAuthentication
    {
        public static UserCredential Authenticate(string[] scopes)
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                var credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                credential =
                    GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }
    }
}
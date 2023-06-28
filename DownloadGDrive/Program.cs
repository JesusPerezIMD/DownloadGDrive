using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using GFile = Google.Apis.Drive.v3.Data.File;

class Program
{
    static string[] Scopes = { DriveService.Scope.Drive };
    static string ApplicationName = "Drive API .NET Quickstart";

    static void Main(string[] args)
    {
        Task.Run(async () =>
        {
            var urls = System.IO.File.ReadAllLines("gdrive.txt");
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var downloadFolderPath = $"GoogleDrive_{timestamp}";
            Directory.CreateDirectory(downloadFolderPath);

            UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore("token.json", true));
            }

            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            foreach (var url in urls)
            {
                try
                {
                    var fileId = url.Split(new string[] { "/file/d/", "/view" }, StringSplitOptions.None)[1];
                    var request = service.Files.Get(fileId);
                    var gfile = await request.ExecuteAsync();
                    var stream = new System.IO.MemoryStream();
                    await request.DownloadAsync(stream);
                    var fileName = Regex.Replace(gfile.Name, @"[^\w\d.]", " ");
                    if (System.IO.File.Exists(Path.Combine(downloadFolderPath, fileName)))
                    {
                        fileName = fileName + "_1";
                    }
                    await System.IO.File.WriteAllBytesAsync(Path.Combine(downloadFolderPath, fileName), stream.ToArray());

                    Console.WriteLine($"Downloaded file: {fileName}");  // Imprimir el mensaje de descarga

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{url}: Error - {ex.Message}");
                }
            }
        }).GetAwaiter().GetResult();
    }
}

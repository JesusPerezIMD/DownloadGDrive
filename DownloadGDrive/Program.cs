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
using System.Diagnostics;
using Google.Apis.Drive.v3.Data;
using OfficeOpenXml;

class Program
{
    static string[] Scopes = { DriveService.Scope.Drive };
    static string ApplicationName = "Drive API .NET Quickstart";

    static async Task<FileList> FindFiles(DriveService service, string fileName)
    {
        var listRequest = service.Files.List();
        listRequest.Q = $"name='{fileName}'";
        listRequest.Fields = "files(id, name)";
        return await listRequest.ExecuteAsync();
    }

    static readonly Dictionary<string, string> MimeTypes = new Dictionary<string, string>
    {
        { "image/jpeg", ".jpg" },
        { "image/png", ".png" },
        { "application/pdf", ".pdf" },
        { "text/plain", ".txt" },
        { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx" },
        { "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx" },
        { "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx" },
        { "application/vnd.ms-excel", ".xls" },
        { "application/vnd.ms-powerpoint", ".ppt" },
        { "application/msword", ".doc" },
    };

    static async Task DownloadFile(DriveService service, GFile file, string downloadFolderPath, StreamWriter writer)
    {
        var request = service.Files.Get(file.Id);
        var gfile = await request.ExecuteAsync();
        var stream = new System.IO.MemoryStream();
        await request.DownloadAsync(stream);

        var fileName = Regex.Replace(gfile.Name, @"[^\w\d.]", " ");
        if (string.IsNullOrEmpty(Path.GetExtension(fileName)) && MimeTypes.TryGetValue(gfile.MimeType, out var extension))
        {
            fileName += extension;
        }

        if (System.IO.File.Exists(Path.Combine(downloadFolderPath, fileName)))
        {
            fileName = fileName + "_1";
        }
        await System.IO.File.WriteAllBytesAsync(Path.Combine(downloadFolderPath, fileName), stream.ToArray());

        Console.WriteLine($"Downloaded file: {fileName}");
        writer.WriteLine($"{fileName}: DESCARGADO ✓");
    }

    static void Main(string[] args)
    {
        Task.Run(async () =>
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(new FileInfo("Bitácora Archivos Drive.xlsx"));
            var worksheet = package.Workbook.Worksheets[0]; // Assume that the data is in the first sheet
            var filenames = new List<string>();
            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                filenames.Add(worksheet.Cells[row, 3].Text); // Column C
            }

            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var downloadFolderPath = $"GoogleDrive_{timestamp}";
            Directory.CreateDirectory(downloadFolderPath);
            var downloadLogPath = Path.Combine(downloadFolderPath, $"descargas_{timestamp}.txt");

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

            using (var writer = new StreamWriter(downloadLogPath))
            {
                foreach (var filename in filenames)
                {
                    try
                    {
                        var fileList = await FindFiles(service, filename);
                        if (fileList.Files == null || fileList.Files.Count == 0)
                        {
                            var filenameWithoutExtension = Path.GetFileNameWithoutExtension(filename);
                            fileList = await FindFiles(service, filenameWithoutExtension);
                            if (fileList.Files == null || fileList.Files.Count == 0)
                            {
                                writer.WriteLine($"{filename}: No se encontró el archivo");
                                continue;
                            }
                        }

                        foreach (var file in fileList.Files)
                        {
                            await DownloadFile(service, file, downloadFolderPath, writer);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{filename}: Error X - {ex.Message}");
                        writer.WriteLine($"{filename}: Error X - {ex.Message}");
                    }
                }
                Process.Start(new ProcessStartInfo(downloadLogPath) { UseShellExecute = true });
            }
        }).GetAwaiter().GetResult();
    }
}

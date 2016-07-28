using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace UploadToGoogleDrive
{
    class Program
    {
        static string[] Scopes = { DriveService.Scope.Drive, DriveService.Scope.DriveFile };
        static string clientId = Properties.Settings.Default.clientId;
        static string clientSecret = Properties.Settings.Default.clientSecret;
        static string inFolder = Properties.Settings.Default.inFolder;
        static string outFolder = Properties.Settings.Default.outFolder;
        static string backupFolder = Properties.Settings.Default.backupFolder;

        static void Main(string[] args)
        {
            try
            {
                UserCredential credential = GetCredentials();

                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                });


                var folderId = GetFolderId(service);

                var files = Directory.GetFiles(inFolder);
                foreach (var file in files)
                {

                    var filePath = file;
                    UploadFile(service, folderId, filePath);
                    System.IO.File.Move(file, outFolder + "/" + new FileInfo(file).Name);
                    RunVBScript(file);
                }
            }
            catch (Exception ex)
            {
                using (var writer = new StreamWriter("errors.txt"))
                {
                    writer.WriteLine("------------------------------------------------");
                    writer.WriteLine(DateTime.Now.ToLongDateString() + " - " + ex.Message);
                    writer.WriteLine(ex.StackTrace);
                    writer.WriteLine("------------------------------------------------");
                }
            }

        }

        private static void RunVBScript(string fileToProcess)
        {
            Process scriptProc = new Process();
            scriptProc.StartInfo.FileName = @"notepad.exe";
            scriptProc.StartInfo.WorkingDirectory = @"c:\temp\"; //<---very important 
            scriptProc.StartInfo.Arguments = "";
            scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Normal; //prevent console window from popping up
            scriptProc.Start();
            scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exit
            scriptProc.Close();
        }

        private static string GetFolderId(DriveService service)
        {
            string folderId = "";

            FilesResource.ListRequest listRequest = service.Files.List();
            IList<Google.Apis.Drive.v2.Data.File> files = listRequest.Execute().Items.ToList().Where(x => x.MimeType == "application/vnd.google-apps.folder" && x.Title == backupFolder).ToList();

            if (files.Any())
                folderId = files.First().Id;
            else
            {
                Google.Apis.Drive.v2.Data.File file = CreateDirectory(service);
                folderId = file.Id;
            }

            return folderId;
        }

        private static Google.Apis.Drive.v2.Data.File CreateDirectory(DriveService service)
        {
            var body = new Google.Apis.Drive.v2.Data.File();
            body.Title = backupFolder;
            body.Description = "document description";
            body.MimeType = "application/vnd.google-apps.folder";

            // service is an authorized Drive API service instance
            var file = service.Files.Insert(body).Execute();
            return file;
        }

        private static void UploadFile(DriveService service, string folderId, string filePath)
        {
            if (System.IO.File.Exists(filePath))
            {
                Google.Apis.Drive.v2.Data.File body = new Google.Apis.Drive.v2.Data.File();

                body.Title = Path.GetFileName(filePath);
                body.Description = "File uploaded by Diamto Drive Sample";
                body.MimeType = GetMimeType(filePath);
                body.Parents = new List<ParentReference>() { new ParentReference() { Id = folderId } };

                byte[] byteArray = System.IO.File.ReadAllBytes(filePath);
                var stream = new MemoryStream(byteArray);

                FilesResource.InsertMediaUpload request = service.Files.Insert(body, stream, GetMimeType(filePath));
                var upload = request.Upload();
                var response = request.ResponseBody;
            }
        }

        private static UserCredential GetCredentials()
        {
            string[] scopes = Scopes;

            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets { ClientId = clientId, ClientSecret = clientSecret },
                                                                        scopes,
                                                                        Environment.UserName,
                                                                        CancellationToken.None,
                                                                        new FileDataStore("Daimto.GoogleDrive.Auth.Store")).Result;

            return credential;
        }

        private static string GetMimeType(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = System.IO.Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            if (regKey != null && regKey.GetValue("Content Type") != null)
                mimeType = regKey.GetValue("Content Type").ToString();
            return mimeType;
        }
    }
}

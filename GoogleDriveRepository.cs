using Aspose.Cells;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
//using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace Test_API.Models
{
    class GoogleDriveRepository
    {
        public static string[] Scopes = { DriveService.Scope.Drive, DriveService.Scope.DriveFile,DriveService.Scope.DriveAppdata };
        static string ApplicationName = "Google Drive API Test";
     
        public DriveService GetService()
        {
            UserCredential credential;
            string credentials = AppDomain.CurrentDomain.BaseDirectory + "credentials.json";
            using (var stream =
                new FileStream(credentials, FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "token.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "allDrives",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        public IList<Google.Apis.Drive.v3.Data.File> ListDriveFiles()
        {
            var GDriveError = "";

            try
            {
                using (DriveService service = GetService())
                {
                    GDriveError = "ListDriveFiles step 1 : GetService()" + service;
                    // Define parameters of request.
                    FilesResource.ListRequest listRequest = service.Files.List();
                    string folder_id = "Talygen";
                    GDriveError = GDriveError + "ListDriveFiles step 2 : service.Files.List()" + listRequest;

                    listRequest.Q = "mimeType='application/vnd.google-apps.folder'";
                    listRequest.SupportsAllDrives = true;
                    listRequest.IncludeItemsFromAllDrives = true;
                    listRequest.Fields = "files(id, name)";
                    listRequest.Fields = "nextPageToken, files(id, name)";

                    IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;

                    GDriveError = GDriveError + "ListDriveFiles step 2 : listRequest.Execute().Files" + files;

                    return files;
                }
            }
            catch (Exception ex)
            {
                string path = HttpContext.Current.Server.MapPath("~/Content/ErrorLog/ErrorLog.txt");
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    writer.WriteLine(GDriveError);
                    writer.WriteLine("ListDriveFiles Error: " + ex.ToString());
                    writer.Close();
                }

                return null;
            }           
           
        }
        //Download GoogleDrive File
        public string DownloadDriveFile(string Fileid)
        {
            using (DriveService service = GetService()) {
                string FolderPath = AppDomain.CurrentDomain.BaseDirectory+"temp\\";
                FilesResource.GetRequest request = service.Files.Get(Fileid);
                request.SupportsAllDrives = true;
                string Filename = request.Execute().Name;
                if(Filename =="On Call All")
                {
                    Filename += ".xslx";
                }
                string FilePath = Path.Combine(FolderPath, Filename);

                MemoryStream stream = new MemoryStream();
                request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
                  {
                      switch (progress.Status)
                      {
                          case DownloadStatus.Downloading:
                              {
                                  Console.WriteLine(progress.BytesDownloaded);
                                  break;
                              }
                          case DownloadStatus.Completed:
                              {
                                   Console.WriteLine("Download Completed");
                                  SaveStream(stream, FilePath);
                                  break;
                              }
                          case DownloadStatus.Failed:
                              {
                                  Console.WriteLine("Download Failed");
                                  break;
                              }

                      }
                  };
                request.Download(stream);
                return FilePath;
        }
        }
        public void SaveStream(MemoryStream stream,string FilePath)
        {
            using (FileStream file = new FileStream(FilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                stream.WriteTo(file);
            }
        }

        public void UploadDriveFile(string path)
        {
            DriveService service = GetService();
            var fileMetadata = new Google.Apis.Drive.v3.Data.File();
            fileMetadata.Name = Path.GetFileName(path);
            fileMetadata.MimeType = "text/csv";
            FilesResource.CreateMediaUpload request;
            using(var stream = new FileStream(path,FileMode.Open))
            {
                request = service.Files.Create(fileMetadata, stream, "text/csv");
                request.Fields = "id";
                request.Upload();

            }
            var file = request.ResponseBody;
            Console.WriteLine("" + file.Id);
        } 
        public string Getfolderid(string foldername)
        {
           
            IList<Google.Apis.Drive.v3.Data.File> files = ListDriveFiles();
            string folderId = "";
            foreach(var file in files)
            {
                if(file.Name == foldername)
                {
                    folderId = file.Id;
                }
            }
            return folderId;

          
        }
        public string getSubfolderId(string FileId,string subfolder)
        {
            string Subfolderid = "";
            using (DriveService service = GetService())
            {
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.Q = "parents='"+FileId+ "' and mimeType='application/vnd.google-apps.folder'";
                listRequest.Fields = "files(id, name)";
                listRequest.Fields = "nextPageToken, files(id, name)";
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                        .Files;
                if (files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        string filename = file.Name;
                        string filenme = filename.Remove(filename.Length - 1, 1);
                        if (subfolder == filename)
                        {
                            Subfolderid = file.Id;
                        }
                        if(subfolder == "April")
                        {
                            if (filenme == "April")
                            {
                                Subfolderid = file.Id;
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No files found.");
                }
            }
            return Subfolderid;
        }
        public IList<Google.Apis.Drive.v3.Data.File> listFolderFiles(string FolderId)
        {
            GoogleDriveFiles fle = new GoogleDriveFiles();
            using (DriveService service = GetService())
            {
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.Q = "parents='" + FolderId + "'";
                listRequest.SupportsAllDrives = true;
                listRequest.IncludeItemsFromAllDrives = true;
                listRequest.Fields = "files(id, name)";
                listRequest.Fields = "nextPageToken, files(id, name)";
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                        .Files;
                return files;
            }
            
            }
        public GoogleDriveFiles GetFileId(string FileId,string date,string reverseDate)
        {
            GoogleDriveFiles fle = new GoogleDriveFiles();
            using (DriveService service = GetService())
            {
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.Q = "parents='" + FileId + "'";
                listRequest.Fields = "files(id, name)";
                listRequest.Fields = "nextPageToken, files(id, name)";
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                        .Files;
                if(files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        if(file.Name.Contains(date)||file.Name.Contains(reverseDate))
                        {
                            fle.Id= file.Id;
                            fle.Name = file.Name;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No files found.");
                }
            }
            return fle;
        }
        public GoogleDriveFiles GetFileId(string FileId, string Filename)
        {
            GoogleDriveFiles fle = new GoogleDriveFiles();
            using (DriveService service = GetService())
            {
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.Q = "parents='" + FileId + "'";
                listRequest.Fields = "files(id, name)";
                listRequest.Fields = "nextPageToken, files(id, name)";
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                        .Files;
                if (files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        if (file.Name == Filename)
                        {
                            fle.Id = file.Id;
                            fle.Name = file.Name;
                        }
                    }
                }
            }
            return fle;
        }



        public List<GoogleDriveFiles> ListFileId()
        {
            using (DriveService service = GetService())
            {
                // Define parameters of request.
                FilesResource.ListRequest listRequest = service.Files.List();
                listRequest.PageSize = 10;
                listRequest.Fields = "nextPageToken, files(id, name)";
                
                // List files.
                IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                    .Files;
                
                List<GoogleDriveFiles> FileList = new List<GoogleDriveFiles>();

                if (files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        GoogleDriveFiles File = new GoogleDriveFiles
                        {
                            Id = file.Id,
                            Name = file.Name,
                            
                    };
                        FileList.Add(File);
                    }
                }
                return FileList;
            }

        }

        // Create Folder in root
        public string CreateFolder(string FolderName)
        {
            IList<Google.Apis.Drive.v3.Data.File> files = ListDriveFiles();
            using (DriveService service = GetService())
            {
                string folderid="";
                var FileMetaData = new Google.Apis.Drive.v3.Data.File();
                bool folderexist = false;
                FileMetaData.Name = FolderName;
                //this mimetype specify that we need folder in google drive
                FileMetaData.MimeType = "application/vnd.google-apps.folder and trashed='false'";
                FilesResource.CreateRequest request;
                request = service.Files.Create(FileMetaData);
                request.Fields = "id";
                foreach(var folder in files)
                {
                    if(folder.Name ==FolderName)
                    {
                        folderexist = true;
                    }

                }
                if(!folderexist)
                {
                    var file = request.Execute();
                    folderid = file.Id.ToString();
                }
                return folderid;
            }
            
        }

        //Upload File
        public string FileUploadInFolder(string foldername, string path)
        {
                //create service
                DriveService service = GetService();
            DeleteFile del = new DeleteFile();
                //get file path
                string filename = Path.GetFileName(path);
            
                string folderId = "";
                IList<Google.Apis.Drive.v3.Data.File> files = ListDriveFiles();
                foreach(var folder in files )
                {
                    if(foldername == folder.Name)
                    {
                        folderId = folder.Id;
                        
                      }
                }
            
            int i = 1;
            IList<Google.Apis.Drive.v3.Data.File> fileslist = listFolderFiles(folderId);
            //List<string> Filelst = new List<string>();
            //foreach(var fle in fileslist)
            //{
            //    Fileslst.Add();
            //}
            restart:
            foreach (var fle in fileslist)
            {
                if (filename == fle.Name)
                {
                   
                    if (filename.Length == 20)
                    {
                        filename = filename.Remove(filename.Length - 4, 4);
                        filename += "_1" + ".csv";
                    }
                    else
                    {
                        filename = filename.Remove(filename.Length - 4, 4);
                        filename = filename.Remove(filename.Length - 2, 2);
                        filename += "_" + i.ToString() + ".csv";
                    }
                    
                    i++;
                    goto restart;
                }
            }
            
            System.IO.FileInfo info = new System.IO.FileInfo(path);
            string renamed_path = AppDomain.CurrentDomain.BaseDirectory + "temp\\" + filename;
            System.IO.FileInfo Renameinfo = new System.IO.FileInfo(renamed_path);
            if(Renameinfo.Exists && i>1)
            {
                del.Delete(renamed_path);
            }
            if (info.Exists)
            {
                info.MoveTo(renamed_path);

            }
            //create file metadata
            var FileMetaData = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = filename,
                    MimeType = MimeMapping.GetMimeMapping(renamed_path),
                    //id of parent folder 
                    Parents = new List<string>
                    {
                        folderId
                    }
                };
                FilesResource.CreateMediaUpload request;
           
            //create stream and upload
            using (var stream = new System.IO.FileStream(renamed_path, System.IO.FileMode.Open))
            {
                request = service.Files.Create(FileMetaData, stream, FileMetaData.MimeType);
                request.Fields = "id";
                request.SupportsAllDrives = true;
                    request.Upload();
             
                //GoogleDriveFiles filelist = GetFileId(folderId,"","");

                var file1 = request.ResponseBody;
            }
            return renamed_path;
            }
       
        public List<DisplayExcelModel> GetHoursPtData(string filePath)
        {
            List<DisplayExcelModel> Values = new List<DisplayExcelModel>();
          
            //Read the contents of CSV file.
            string[] csvData = System.IO.File.ReadAllLines(filePath);
            
            for(int i=0;i<csvData.Length;i++)
            {

                string[] nums =csvData[i].Split(',').ToArray();
                if (nums[0] != "Evaluation Only. Created with Aspose.Cells for .NET.Copyright 2003 - 2021 Aspose Pty Ltd."&&nums[0]!="\r"&&nums[0]!="")
                {
                    try
                    {
                        Values.Add(new DisplayExcelModel(nums[0], nums[1], nums[2], nums[3], nums[4], nums[5], nums[6], nums[7], nums[8], nums[9], nums[10], nums[11]));
                    }
                    catch(Exception e)
                    {

                    }

                   
                }
            }
           
            return Values;
        }
     

        public string trimcsv(string Filepath,string date)
        {

                 
                string[] values = System.IO.File.ReadAllLines(Filepath);
            DeleteFile del = new DeleteFile();
            del.Delete(Filepath);
            string auditPath = AppDomain.CurrentDomain.BaseDirectory + "temp\\" + "Audit_"+date+".csv";
                StreamWriter Writer = new StreamWriter(auditPath, false);
                for (int i = 0; i < values.Length; i++)
                {//replacing null with " " (space)
                    if(values[i].Contains("Evaluation Only. Created with Aspose.Cells for .NET.Copyright 2003 - 2021 Aspose Pty Ltd"))
                {
                    string input = values[i].Split(',').First();
                    Writer.WriteLine(values[i].Replace(input, " "));
                    
                }
                else
                {
                    Writer.WriteLine(values[i]);
                }
                    
                }

                Writer.Close();
            return auditPath;
            }
        
    }

    }
        

    

    

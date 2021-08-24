using System;
using System.IO;
using System.Linq;
using System.Web;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using System.Security;
using ClientOM = Microsoft.SharePoint.Client;
using System.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using Robotics.Model.Job.CodeF;
using Robotics.Services.Data;
using System.Threading.Tasks;

namespace SharepointDownloader
{
    public class SharepointDownloader
    {

        public ClientContext clientContext { get; set; }
       
        private Web WebClient { get; set; }

        public SharepointDownloader(string ServerSiteUrl,string Password, string UserName)
        {
            this.Connect(ServerSiteUrl, Password, UserName);
        }

        public void Connect(string ServerSiteUrl, string Password, string UserName)
        {
            try
            {
                using (clientContext = new ClientContext(ServerSiteUrl))
                {
                    var securePassword = new SecureString();
                    foreach (char c in Password)
                    {
                        securePassword.AppendChar(c);
                    }

                    clientContext.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                    WebClient = clientContext.Web;
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        public string DownloadFiles(string file_directory,string LocalPath)//link,filename,folder,local path
        {
            try
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(LocalPath);
                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
                FileCollection files = WebClient.GetFolderByServerRelativeUrl(file_directory).Files;
                clientContext.Load(files);
                clientContext.ExecuteQuery();

                if (clientContext.HasPendingRequest)
                    clientContext.ExecuteQuery();

                foreach (ClientOM.File file in files)
                {
                    FileInformation fileInfo = ClientOM.File.OpenBinaryDirect(clientContext, file.ServerRelativeUrl);
                    clientContext.ExecuteQuery();

                    var filePath = LocalPath + file.Name;
                    using (var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
                    {
                        fileInfo.Stream.CopyTo(fileStream);
                    }
                }
               
                //files.ToList().ForEach(file => file.DeleteObject());
                //clientContext.ExecuteQuery();
                //Console.WriteLine("File successfully deleted from source location!");

                return "";
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }
        public void UploadFileToSharepoint(string localFilePath,string libraryname,string serverfoldername) //link,filename,folder
        {
            string serverRelUrl = "";
            System.IO.FileStream fs = null;
            try
            {
                List myLibrary = WebClient.Lists.GetByTitle(libraryname);
                Folder folder = myLibrary.RootFolder;
                clientContext.Load(folder, f => f.ServerRelativeUrl);
                clientContext.ExecuteQuery();
                serverRelUrl = folder.ServerRelativeUrl;
                serverRelUrl = serverRelUrl + serverfoldername;
                fs = new System.IO.FileStream(localFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
               Microsoft.SharePoint.Client.File.SaveBinaryDirect(clientContext, serverRelUrl + "/" + System.IO.Path.GetFileName(localFilePath), fs, true);
               Console.WriteLine("File successfully uploaded to new path");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                fs.Dispose();
            }
        }
    }
        }
    


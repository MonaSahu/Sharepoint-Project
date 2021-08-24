using System;
using Robotics.Services.Configuration;
using System.Configuration;
using System.IO;

namespace SharepointDownloader
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
       
         
        static void Main(string[] args)

        {
            Robotics.Services.Configuration.ConfigManager.LoadConfigFromObjectStore().Wait();
            string outputfolder = ConfigManager.Settings.SolutionProperties.Output; // for getting from JSON
            string ServerSiteUrl = ConfigManager.Settings.SolutionProperties.Sharepoint_url;
            string UserName =ConfigurationManager.AppSettings["UserName"];//email id used to login to sharepoint
            string Password =ConfigurationManager.AppSettings["Password"]; //password
            string sharedrive_location = ConfigManager.Settings.SolutionProperties.Sharepoint_RelativePath;
            string serverfoldername = ConfigurationManager.AppSettings["UploadDirectory"];
            string libraryname = ConfigManager.Settings.SolutionProperties.Library;
            string localPath = ConfigurationManager.AppSettings["Local_Dir"];
            string uploadPath = ConfigurationManager.AppSettings["Local"];

            SharepointDownloader sharepointModel = new SharepointDownloader(ServerSiteUrl, UserName, Password);
            try
            {
                    sharepointModel.DownloadFiles(sharedrive_location, localPath);
                    DirectoryInfo directory = new DirectoryInfo(localPath); //Config File which is individual Customer
                    FileInfo[] files = directory.GetFiles("*.*");
                    foreach (FileInfo file in files)
                    {
                        var filePath = localPath + file.Name;
                        sharepointModel.UploadFileToSharepoint(filePath, libraryname,serverfoldername);
                        string ArcMove = localPath + "/Archive/";
                        string stampedFileName = file.Name.Replace(".",string.Format("{0:YYYY-mm-dd hhmmss}", DateTime.UtcNow) + ".");
                        ArcMove = ArcMove + stampedFileName;
                        file.MoveTo(ArcMove);
                    }
                      }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
    }



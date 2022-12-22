using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Policy;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace AreaPrice.Scrapping
{
    public partial class AreaPriceScrapping : ServiceBase
    {
        private bool WorkingOnProgress = false;

        /// <summary>
        /// Default constructor
        /// </summary>
        public AreaPriceScrapping()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Override the method called on service started
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            try
            {
                Log($"Service Started");
            }
            catch (Exception ex)
            {

                eventLog.WriteEntry(ex.ToString(), EventLogEntryType.Error);
            }

            try
            {
                var intervalMinutes = GetIntervalMinutes();
                Log($"Get Interval Minutes {intervalMinutes}");

                var timer = new Timer(intervalMinutes * 60 * 1000);
                timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
                timer.Start();
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.ToString(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Override the method called on service stopped
        /// </summary>
        protected override void OnStop()
        {
            Log("Service Stopped");
        }

        /// <summary>
        /// The on action timer method 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnTimer(object sender, ElapsedEventArgs args)
        {
            Log("Start Scrapping");

            try
            {
                if (WorkingOnProgress)
                    return;

                WorkingOnProgress = true;

                var uriDAM = "https://www.iexindia.com/marketdata/areaprice.aspx";
                DownloadFile(uriDAM, FileType.DAM);

                var uriRTM = "https://www.iexindia.com/marketdata/rtm_areaprice.aspx";
                DownloadFile(uriRTM, FileType.RTM);
            }
            catch (Exception ex)
            {
                Log(ex);
            }
            finally
            {
                WorkingOnProgress = false;
            }
        }

        /// <summary>
        /// Download excel file 
        /// </summary>
        /// <returns></returns>
        private void DownloadFile(string baseUri, FileType fileType)
        {
            try
            {
                Log($"Start Download Excel File {fileType}");

                var cookieContainer = new CookieContainer();

                using (var httpClientHandler = new HttpClientHandler
                {
                    CookieContainer = cookieContainer
                })
                {
                    Log($"Start get session data {fileType}");

                    var uri = new Uri(baseUri);
                    using (var httpClient = new HttpClient(httpClientHandler))
                    {
                        var response = httpClient.GetAsync(uri).Result;
                        var htmlContent = response.Content.ReadAsStringAsync().Result;

                        Log($"Get HTMl content {fileType}");

                        var htmlDocument = new HtmlDocument();
                        htmlDocument.LoadHtml(htmlContent);

                        // Extract the control id from HTML content
                        var reportViewerElement = htmlDocument.GetElementbyId("ctl00_InnerContent_reportViewer_ctl09_ReportControl");
                        var childElement = reportViewerElement.ChildNodes[5].ChildNodes[1];

                        // Parse control id by removing extra data 
                        var controlId = childElement.Id.Replace("_1_oReportDiv", "").Remove(0, 1);

                        Log($"Get Control Id {controlId} , File Type {fileType}");

                        // Retrive the cookie from the header
                        var cookie = cookieContainer.GetCookieHeader(uri).ToString();

                        Log($"Get Cookie {cookie} , File Type {fileType}");

                        var fileUri = $"https://www.iexindia.com/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID={controlId}&Mode=true&OpType=Export&FileName=PriceMinute&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML";
                        httpClient.DefaultRequestHeaders.Add("Cookie", cookie);

                        var filePath = Path.Combine(GetExportFolderPath(fileType), GetFileName(fileType));

                        Log($"Download excel file. URI: {uri} , FilePath: {filePath} , File Type {fileType}");

                        ArchiveExportedFiles(fileType);

                        byte[] fileBytes = httpClient.GetByteArrayAsync(fileUri).Result;
                        File.WriteAllBytes(filePath, fileBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }

        /// <summary>
        /// Get the interval minutes from the configurations
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private int GetIntervalMinutes()
        {
            const string IntervalMinutesKeyName = "IntervalMinutes";
            const int DefaultIntervalMinutes = 1; // Default Interval 60 Minutes

            var intervalMinutesValue = ConfigurationManager.AppSettings[IntervalMinutesKeyName];
            if (intervalMinutesValue == null)
                return DefaultIntervalMinutes;

            if (int.TryParse(intervalMinutesValue, out var intervalMinutes))
                return intervalMinutes;

            return DefaultIntervalMinutes;
        }

        /// <summary>
        /// Get the export folder path
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private string GetExportFolderPath(FileType fileType)
        {
            string folderPathKeyName = $"ExportFolderPath-{fileType}";

            var exportFolderPath = ConfigurationManager.AppSettings[folderPathKeyName];
            if (string.IsNullOrEmpty(exportFolderPath) == false)
                return exportFolderPath;

            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var exportDir = Path.Combine(baseDir, "Exports", fileType.ToString());

            if (Directory.Exists(exportDir) == false)
                Directory.CreateDirectory(exportDir);

            return exportDir;
        }

        /// <summary>
        /// Get the export folder path
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private void ArchiveExportedFiles(FileType fileType)
        {
            var folderPath = GetExportFolderPath(fileType);
            var archiveFolderPath = Path.Combine(folderPath, "Archive");

            if (Directory.Exists(archiveFolderPath) == false)
                Directory.CreateDirectory(archiveFolderPath);

            var files = Directory.GetFiles(folderPath);
            foreach (var file in files)
            {
                try
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.Extension != ".xlsx")
                        continue;

                    var archiveFilePath = Path.Combine(archiveFolderPath, fileInfo.Name);
                    if (File.Exists(archiveFilePath) == false)
                        File.Copy(file, archiveFilePath);

                    File.Delete(file);
                }
                catch (Exception ex)
                {
                    Log(ex);
                }
            }
        }

        /// <summary>
        /// Get the file name according to its type
        /// </summary>
        /// <param name="fileType"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private string GetFileName(FileType fileType)
        {
            switch(fileType) 
            {
                case FileType.DAM:
                    return $"{DateTime.Now:yyyyMMdd_DAM_HH.mm}.xlsx";

                case FileType.RTM:
                    return $"{DateTime.Now:yyyyMMdd}_RTM1hr_{DateTime.Now:HH.mm}.xlsx";

                default:
                    throw new NotImplementedException();
            }
        }
        /// <summary>
        /// Log Message
        /// </summary>
        /// <param name="message"></param>
        private void Log(string message)
        {
            try
            {
                var logFile = Path.Combine(GetLogDirectory(), $"log_{DateTime.Now:yyyy-MM-dd}.txt");

                if (File.Exists(logFile) == false)
                    File.Create(logFile).Close();

                eventLog.WriteEntry(message, EventLogEntryType.Information);

                using (StreamWriter w = File.AppendText(logFile))
                {
                    Log(message, w);
                }
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.ToString(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Log Message
        /// </summary>
        /// <param name="message"></param>
        private void Log(Exception exception)
        {
            try
            {
                var logFile = Path.Combine(GetLogDirectory(), $"log-errors_{DateTime.Now:yyyy-MM-dd}.txt");

                if (File.Exists(logFile) == false)
                    File.Create(logFile).Close();

                eventLog.WriteEntry(exception.ToString(), EventLogEntryType.Error);

                using (StreamWriter w = File.AppendText(logFile))
                {
                    Log(exception.ToString(), w);
                }
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.ToString(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="logMessage"></param>
        /// <param name="w"></param>
        public static void Log(string logMessage, TextWriter w)
        {
            w.Write("\r\nLog Entry : ");
            w.WriteLine($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
            w.WriteLine($"  :{logMessage}");
            w.WriteLine("-------------------------------");
        }

        /// <summary>
        /// Get or create log directory
        /// </summary>
        /// <returns></returns>
        private static string GetLogDirectory()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var logDir = Path.Combine(baseDir, "logs");
            
            if (Directory.Exists(logDir) == false)
                Directory.CreateDirectory(logDir);

            return logDir;
        }

        /// <summary>
        /// File Type
        /// </summary>
        private enum FileType
        {
            DAM,
            RTM
        }
    }
}

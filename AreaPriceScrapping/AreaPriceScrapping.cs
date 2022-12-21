﻿using HtmlAgilityPack;
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
            eventLog.WriteEntry("Started.");

            try
            {
                var intervalMinutes = GetIntervalMinutes();

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
            eventLog.WriteEntry("Stopped");
        }

        /// <summary>
        /// The on action timer method 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnTimer(object sender, ElapsedEventArgs args)
        {
            eventLog.WriteEntry("Start Scrapping", EventLogEntryType.Information);

            try
            {
                DownloadExcelFile();
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.ToString(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Download excel file 
        /// </summary>
        private void DownloadExcelFile()
        {
            var (cookie, controlId) = GetSessionData();
            using (var webClient = new WebClient())
            {
                var uri = $"https://www.iexindia.com/Reserved.ReportViewerWebControl.axd?Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID={controlId}&Mode=true&OpType=Export&FileName=PriceMinute&ContentDisposition=OnlyHtmlInline&Format=EXCELOPENXML";
                webClient.Headers.Add("Cookie", cookie);

                var fileName = $"{DateTime.Now.ToString("yyyy-MM-dd HH.mm")}.xlsx";
                var filePath = Path.Combine(GetExportFolderPath(), fileName);
                
                webClient.DownloadFile(uri, filePath);
            }
        }

        /// <summary>
        /// Get session data required to download excel file
        /// </summary>
        /// <returns></returns>
        private (string cookie, string controlId) GetSessionData()
        {
            var baseUri = "https://www.iexindia.com/marketdata/areaprice.aspx";

            var cookie = string.Empty;
            var controlId = string.Empty;

            var cookieContainer = new CookieContainer();

            using (var httpClientHandler = new HttpClientHandler
            {
                CookieContainer = cookieContainer
            })
            {
                var uri = new Uri(baseUri);
                using (var httpClient = new HttpClient(httpClientHandler))
                {
                    var response = httpClient.GetAsync(uri).Result;
                    var htmlContent = response.Content.ReadAsStringAsync().Result;

                    var htmlDocument = new HtmlDocument();
                    htmlDocument.LoadHtml(htmlContent);

                    // Extract the control id from HTML content
                    var reportViewerElement = htmlDocument.GetElementbyId("ctl00_InnerContent_reportViewer_ctl09_ReportControl");
                    var childElement = reportViewerElement.ChildNodes[5].ChildNodes[1];

                    // Parse control id by removing extra data 
                    controlId = childElement.Id.Replace("_1_oReportDiv", "").Remove(0, 1);

                    // Retrive the cookie from the header
                    cookie = cookieContainer.GetCookieHeader(uri).ToString();
                }
            }

            return (cookie, controlId);
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
        private string GetExportFolderPath()
        {
            const string FolderPathKeyName = "ExportFolderPath";

            var exportFolderPath = ConfigurationManager.AppSettings[FolderPathKeyName];
            if (exportFolderPath == null)
                return string.Empty;

            return exportFolderPath;
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Common;
using Discovery;
using System.Net;
using System.Security;
using WorkflowScanner.Models;
using System.Collections.Generic;
using System.IO;
using Common.Models;

namespace ScannerTest
{
    [TestClass]
    public class WorkflowScannerTests
    {
        [TestMethod]
        public void SingleSiteCollection()
        {
            string AssessmentOutputFolder = System.Configuration.ConfigurationManager.AppSettings["AssessmentOutputFolder"];
            string logFolderPath;
            DirectoryInfo logFolder;
            if (!Directory.Exists(string.Concat(AssessmentOutputFolder, @"\Logs")))
            {
                logFolder = System.IO.Directory.CreateDirectory(string.Concat(AssessmentOutputFolder, @"\Logs"));
                logFolderPath = logFolder.FullName;
                Logging.LOG_DIRECTORY = logFolderPath;
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Log folder created");

            }
            else
            {
                logFolderPath = string.Concat(AssessmentOutputFolder, @"\Logs");
                Logging.LOG_DIRECTORY = logFolderPath;
            }

            ReportGeneration ops = new ReportGeneration();
            WorkflowScanning objonPrem = new WorkflowScanning();

            ops.CreateDirectoryStructure(AssessmentOutputFolder);

            objonPrem.Scope = AnalysisScope.SiteCollection;
            objonPrem.Url = System.Configuration.ConfigurationManager.AppSettings["TestSiteCollectionURL"];
            objonPrem.Credential = new NetworkCredential(System.Configuration.ConfigurationManager.AppSettings["TestUsername"], 
                System.Configuration.ConfigurationManager.AppSettings["TestPassword"], 
                System.Configuration.ConfigurationManager.AppSettings["TestDomain"]);

            List<WorkflowScanResult> results = objonPrem.Scan();

            Assert.IsTrue(results.Count == Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SingleSiteCollectionExpectedCount"]));
        }

        [TestMethod]
        public void WebApp()
        {
            string AssessmentOutputFolder = System.Configuration.ConfigurationManager.AppSettings["AssessmentOutputFolder"];
            string logFolderPath;
            DirectoryInfo logFolder;
            if (!Directory.Exists(string.Concat(AssessmentOutputFolder, @"\Logs")))
            {
                logFolder = System.IO.Directory.CreateDirectory(string.Concat(AssessmentOutputFolder, @"\Logs"));
                logFolderPath = logFolder.FullName;
                Logging.LOG_DIRECTORY = logFolderPath;
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Log folder created");

            }
            else
            {
                logFolderPath = string.Concat(AssessmentOutputFolder, @"\Logs");
                Logging.LOG_DIRECTORY = logFolderPath;
            }

            ReportGeneration ops = new ReportGeneration();
            WorkflowScanning objonPrem = new WorkflowScanning();

            ops.CreateDirectoryStructure(AssessmentOutputFolder);

            objonPrem.Scope = AnalysisScope.WebApplication;
            objonPrem.Url = System.Configuration.ConfigurationManager.AppSettings["TestWebApplicationURL"];

            List<WorkflowScanResult> results = objonPrem.Scan();

            Assert.IsTrue(results.Count == Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["WebAppExpectedCount"]));
        }

        [TestMethod]
        public void Farm()
        {
            string AssessmentOutputFolder = System.Configuration.ConfigurationManager.AppSettings["AssessmentOutputFolder"];
            string logFolderPath;
            DirectoryInfo logFolder;
            if (!Directory.Exists(string.Concat(AssessmentOutputFolder, @"\Logs")))
            {
                logFolder = System.IO.Directory.CreateDirectory(string.Concat(AssessmentOutputFolder, @"\Logs"));
                logFolderPath = logFolder.FullName;
                Logging.LOG_DIRECTORY = logFolderPath;
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Log folder created");

            }
            else
            {
                logFolderPath = string.Concat(AssessmentOutputFolder, @"\Logs");
                Logging.LOG_DIRECTORY = logFolderPath;
            }

            ReportGeneration ops = new ReportGeneration();
            WorkflowScanning objonPrem = new WorkflowScanning();

            ops.CreateDirectoryStructure(AssessmentOutputFolder);

            objonPrem.Scope = AnalysisScope.Farm;

            List<WorkflowScanResult> results = objonPrem.Scan();

            Assert.IsTrue(results.Count == Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["FarmExpectedCount"]));
        }
    }
}
﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Common;
using Discovery;
using System.Net;
using System.Security;
using WorkflowScanner.Models;
using System.Collections.Generic;
using System.IO;
using Common.Models;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using PowerShell;
using Test;
using static System.Net.WebRequestMethods;
using System.Linq;
using System.Globalization;

namespace ScannerTest
{
    [TestClass]
    public class PowerShellTests
    {
        private Runspace _runspace;

        [TestInitialize]
        public void Init()
        {
            var initialSessionState = InitialSessionState.CreateDefault();
            initialSessionState.Commands.Add(
                new SessionStateCmdletEntry("Get-SPWorkflowScannerWorkflows", typeof(CmdGetWorkflowAssociationsForOnprem), null)
            );
            _runspace = RunspaceFactory.CreateRunspace(initialSessionState);

            _runspace.Open();
        }

        [TestMethod]
        public void TestPowerShellSingleSiteCollection()
        {
            using (var powershell = System.Management.Automation.PowerShell.Create())
            {
                string AssessmentOutputFolder = System.Configuration.ConfigurationManager.AppSettings["AssessmentOutputFolder"];

                powershell.Runspace = _runspace;

                var command = new Command("Get-SPWorkflowScannerWorkflows");
                command.Parameters.Add("AssessmentOutputFolder", AssessmentOutputFolder);
                command.Parameters.Add("SiteCollectionUrl", System.Configuration.ConfigurationManager.AppSettings["TestSiteCollectionURL"]);
                command.Parameters.Add("DomainName", System.Configuration.ConfigurationManager.AppSettings["TestDomain"]);
                command.Parameters.Add("Credential", new PSCredential(System.Configuration.ConfigurationManager.AppSettings["TestUsername"],
                                      System.Configuration.ConfigurationManager.AppSettings["TestPassword"].ToSecureString()));
                command.Parameters.Add("Force");
                powershell.Commands.AddCommand(command);

                powershell.Invoke();

                // Get rows in resulting file
                string csvFilePath = GetLatestCsvFilePath($"{AssessmentOutputFolder}\\Summary");
                int rows = CountRowsInCsvFile(csvFilePath);

                Assert.IsTrue(rows == Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SingleSiteCollectionExpectedCount"]));
            }
        }

        private int CountRowsInCsvFile(string filePath)
        {
            return System.IO.File.ReadLines(filePath).Count() - 1; // Counts lines excluding the header
        }

        private string GetLatestCsvFilePath(string folderPath)
        {
            DirectoryInfo directory = new DirectoryInfo(folderPath);
            FileInfo latestFile = directory.GetFiles("*.csv") // Gets all csv files
                    .OrderByDescending(f => f.LastWriteTime) // Orders them by creation time
                    .First(); // Gets the newest file

            return latestFile.FullName;
        }
    }
}
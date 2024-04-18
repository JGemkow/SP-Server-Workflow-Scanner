using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Management.Automation;
using Common;
using Discovery;
using System.Net;
using Common.Models;


namespace Root
{
    /// <summary>
    /// Commandlet to discover the Workflow Asssociations in 
    /// Onpremise environments that are published to
    /// list and document libraries
    /// </summary>
    /// <returns></returns>
    ///

    [Cmdlet(VerbsCommon.Get, "WorkflowAssociationsForOnprem", DefaultParameterSetName = "CurrentCredential-Farm")]
    public class CmdGetWorkflowAssociationsForOnprem : PSCmdlet
    {
        [Parameter(Mandatory = true, ParameterSetName = "Credential-WebApplication", HelpMessage = "Specify the Domain name of the user account")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollection", HelpMessage = "Specify the Domain name of the user account")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollectionUrls", HelpMessage = "Specify the Domain name of the user account.")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-Farm", HelpMessage = "Specify the Domain name of the user account")]
        public string DomainName;

        [Parameter(Mandatory = true, ParameterSetName = "Credential-WebApplication", HelpMessage = "Specify the user name and password. If not specified, current credentials will be used.")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollection", HelpMessage = "Specify the user name and password. If not specified, current credentials will be used.")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollectionUrls", HelpMessage = "Specify the user name and password. If not specified, current credentials will be used.")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-Farm", HelpMessage = "Specify the user name and password. If not specified, current credentials will be used.")]
        public PSCredential Credential;

        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollectionUrls", HelpMessage = "Specify the file path of a text file containing target site collection URLs")]
        [Parameter(Mandatory = true, ParameterSetName = "CurrentCredential-SiteCollectionUrls", HelpMessage = "Specify the file path of a text file containing target site collection URLs")]
        public string SiteCollectionURLFilePath;

       [Parameter(Mandatory = true, ParameterSetName = "Credential-WebApplication", HelpMessage = "Specify the URL of the Web Application")]
       [Parameter(Mandatory = true, ParameterSetName = "CurrentCredential-WebApplication", HelpMessage = "Specify the URL of the Web Application")]
        public string WebApplicationUrl;

        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollection", HelpMessage = "Specify the URL of the Site Collection")]
        [Parameter(Mandatory = true, ParameterSetName = "CurrentCredential-SiteCollection", HelpMessage = "Specify the URL of the Site Collection")]
        public string SiteCollectionUrl;

        [Parameter(Mandatory = true, ParameterSetName = "Credential-WebApplication", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollection", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-SiteCollectionUrls", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        [Parameter(Mandatory = true, ParameterSetName = "Credential-Farm", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        [Parameter(Mandatory = true, ParameterSetName = "CurrentCredential-WebApplication", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        [Parameter(Mandatory = true, ParameterSetName = "CurrentCredential-SiteCollection", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        [Parameter(Mandatory = true, ParameterSetName = "CurrentCredential-SiteCollectionUrls", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        [Parameter(Mandatory = true, ParameterSetName = "CurrentCredential-Farm", HelpMessage = @"The path where the Assessment Summary, logs, Workflow definitions are downloaded (if DownloadForms parameter is set to true) for analyzing (e.g. F:\temp\WorkflowDefinitions")]
        public string AssessmentOutputFolder;

        private string assessmentScope;
        private static List<string> siteCollectionUrls = new List<string>();
        private string logFolderPath;
        private DirectoryInfo logFolder;
        private DataTable dtWorkflowLocations = new DataTable();

        protected override void BeginProcessing()
        {

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

            base.BeginProcessing();
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

        protected override void ProcessRecord()
        {
            try
            {
                assessmentScope = ParameterSetName.Substring(ParameterSetName.IndexOf('-'), ParameterSetName.Length - ParameterSetName.IndexOf('-'));
                BeginToAssess();
            }
            catch (Exception ex)
            {
                Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
            }
        }


        protected void BeginToAssess()
        {
            ReportGeneration ops = new ReportGeneration();
            try
            {
                //New Code Starts
                string userInput = string.Empty;
                Console.WriteLine(System.Environment.NewLine);
                Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "The assessment is scoped to run at " + assessmentScope +
                       " level. Would you like to proceed? [Y] to continue, [N] to abort.");
                var op = this.InvokeCommand.InvokeScript("Read-Host");
                userInput = op[0].ToString().Trim().ToLower();

                while (!userInput.Equals("y") && !userInput.Equals("n"))
                {
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Invalid input. Press [Y] to continue, [N] to abort.");
                    op = this.InvokeCommand.InvokeScript("Read-Host");
                    userInput = op[0].ToString().ToLower();

                }
                if (userInput.Equals("y"))
                {
                    WorkflowScanning objonPrem = new WorkflowScanning();
                    ops.CreateDirectoryStructure(AssessmentOutputFolder);
                    Console.WriteLine(System.Environment.NewLine);
                    Host.UI.WriteLine(ConsoleColor.Yellow, Host.UI.RawUI.BackgroundColor, "Beginning assessment..");

                    switch (assessmentScope)
                    {
                        case "Farm":
                            {
                                objonPrem.Scope = AnalysisScope.Farm;
                                objonPrem.Url = null;
                                break;
                            }
                        case "WebApplication":
                            {
                                objonPrem.Scope = AnalysisScope.WebApplication;
                                objonPrem.Url = WebApplicationUrl;
                                break;
                            }
                        case "SiteCollection":
                            {
                                objonPrem.Scope = AnalysisScope.SiteCollection;
                                objonPrem.Url = SiteCollectionUrl;
                                break;
                            }
                        case "SiteCollectionsUrls":
                            {
                                objonPrem.Scope = AnalysisScope.SiteCollectionsUrls;
                                objonPrem.Url = SiteCollectionURLFilePath;
                                break;
                            }
                    }

                    objonPrem.DownloadPath = AssessmentOutputFolder;

                    //Set Credentials from user entry if provided
                    if (Credential != null)
                    {
                        objonPrem.Credential = new NetworkCredential(Credential.UserName, Credential.Password.ToString(), DomainName);
                    }
                    
                    // run the workflow scan
                    dtWorkflowLocations = objonPrem.Scan().ToDataTable();

                    //Save the CSV file
                    string csvFilePath = string.Concat(AssessmentOutputFolder, ops.summaryFolder, ops.summaryFile);
                    ops.WriteToCsvFile(dtWorkflowLocations, csvFilePath);
                }
                else if (userInput.Equals("n"))
                {
                    Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Operation aborted as per your input !");
                }              
            }
            catch (Exception ex)
            {
                Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
            }
        }
    }
}

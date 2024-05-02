using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Management.Automation;
using Common;
using Discovery;
using System.Net;
using Common.Models;


namespace PowerShell
{
    /// <summary>
    ///   Commandlet to discover the Workflow Asssociations in 
    ///   on-premises environments that are published to
    ///   list and document libraries
    /// </summary
    /// <example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -SiteCollectionURLFilePath .\sites.csv</code>
    /// <para>Using credentials of current logged in user, scan all sites in the sites.csv.</para>
    /// </example>
    /// <example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -SiteCollectionURLFilePath .\sites.csv -Credential $Cred -Domain "ad.xyz123domain.com"</code>
    /// <para>Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain,scan all sites in the sites.csv.</para>
    /// </example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -SiteCollectionUrl "https://sp2016.ad.xyz123domain.com/sites/Test123"</code>
    /// <para>Using credentials of current logged in user, scan the single Test123 site collection.</para>
    /// </example
    /// <example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -SiteCollectionUrl "https://sp2016.ad.xyz123domain.com/sites/Test123" -Credential $Cred -Domain "ad.xyz123domain.com"</code>
    /// <para>Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain,scan the single Test123 site collection.</para>
    /// </example>
    /// <example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -WebApplicationUrl "https://sp2016.ad.xyz123domain.com"</code>
    /// <para>Using a PSCredential object ($Cred), enumerate all sites in the provided web application and scan for workflow in the site collection.</para>
    /// </example>
    /// /// <example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -WebApplicationUrl "https://sp2016.ad.xyz123domain.com" -Credential $Cred -Domain "ad.xyz123domain.com"</code>
    /// <para>Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain, enumerate all sites in the provided web application and scan for workflow in the site collection.</para>
    /// </example>
    /// <example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput"</code>
    /// <para>Using credentials of current logged in user, enumerate all sites in the farm. NOTE: This must be run on a SharePoint Server.</para>
    /// </example>
    /// <example>
    /// <code>Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -Credential $Cred -Domain "ad.xyz123domain.com"</code>
    /// <para>Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain, enumerate all sites in the farm. NOTE: This must be run on a SharePoint Server.</para>
    /// </example>
    /// <param name="DomainName">Name of the domain paired with a provided credential.</param>
    /// <param name="Credential">Credential object used for authentication against SharePoint.</param>
    /// <param name="SiteCollectionURLFilePath">Path to file containing URLs of site collections to be scanned.</param>
    /// <param name="WebApplicationUrl">URL of the web application in SharePoint for sites to be enumerated and scanned.</param>
    /// <param name="SiteCollectionUrl">URL of a single site collection in SharePoint to be scanned.</param>
    /// <param name="AssessmentOutputFolder">Path to folder where assessment output should be stored.</param>
    /// <output>Produces a WorkflowSummary.csv file listing detected workflows.</output>
    [Cmdlet(VerbsCommon.Get, "SPWorkflowScannerWorkflows", DefaultParameterSetName = "CurrentCredential-Farm")]
    public class CmdGetSPWorkflowScannerWorkflows : PSCmdlet
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

        [Parameter(Mandatory = false)]
        public SwitchParameter Force
        {
            get { return _force; }
            set { _force = value; }
        }
        private bool _force;

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
                assessmentScope = ParameterSetName.Substring(ParameterSetName.IndexOf('-')+1, ParameterSetName.Length - ParameterSetName.IndexOf('-') -1);
                BeginToAssess();
            }
            catch (Exception ex)
            {
                try
                {
                    Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
                }
                catch (System.Management.Automation.Host.HostException)
                {
                    Console.WriteLine(ex.Message);
                }
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
                if (Force.IsPresent)
                {
                    userInput = "y";
                }
                else
                {
                    while (!userInput.Equals("y") && !userInput.Equals("n"))
                    {
                        Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Invalid input. Press [Y] to continue, [N] to abort.");
                        var op = this.InvokeCommand.InvokeScript("Read-Host");
                        userInput = op[0].ToString().ToLower();

                    }
                } 

                if (userInput.Equals("y"))
                {
                    WorkflowScanning objonPrem = new WorkflowScanning();
                    ops.CreateDirectoryStructure(AssessmentOutputFolder);
                    Console.WriteLine(System.Environment.NewLine);
                    try
                    {
                        Host.UI.WriteLine(ConsoleColor.Yellow, Host.UI.RawUI.BackgroundColor, "Beginning assessment..");
                    }
                    catch (System.Management.Automation.Host.HostException)
                    {
                        Console.WriteLine("Beginning assessment..");
                    }

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
                        objonPrem.Credential = Credential.GetNetworkCredential();
                        objonPrem.Credential.Domain = DomainName;
                    }
                    
                    // run the workflow scan
                    dtWorkflowLocations = objonPrem.Scan().ToDataTable();

                    //Save the CSV file
                    string csvFilePath = string.Concat(AssessmentOutputFolder, ops.summaryFolder, ops.summaryFile);
                    ops.WriteToCsvFile(dtWorkflowLocations, csvFilePath);
                }
                else if (userInput.Equals("n"))
                {
                    try
                    {
                        Host.UI.WriteLine(ConsoleColor.Cyan, Host.UI.RawUI.BackgroundColor, "Operation aborted as per your input !");
                    }
                    catch (System.Management.Automation.Host.HostException)
                    {
                        Console.WriteLine("Operation aborted as per your input !");
                    }
                }              
            }
            catch (Exception ex)
            {
                try
                {
                    Host.UI.WriteLine(ConsoleColor.DarkRed, Host.UI.RawUI.BackgroundColor, ex.Message);
                }
                catch (System.Management.Automation.Host.HostException)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}

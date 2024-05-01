using Common;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using WorkflowScanner.Models;
using System.Linq;
using Common.Models;


namespace Discovery
{
    public class WorkflowScanning
    {
        public string Url { get; set; }
        public string FilePath { get; set; }
        
        public AnalysisScope Scope { get; set; }
        public bool OnPrem { get; set; }
        public string DownloadPath { get; set; }
        public NetworkCredential Credential {get;set;}

        public DirectoryInfo analysisFolder;
        public DirectoryInfo downloadedFormsFolder;
        public DirectoryInfo summaryFolder;

        public DataTable dt = new DataTable();
        public List<WorkflowScanResult> Results = new List<WorkflowScanResult>();

        public List<WorkflowScanResult> Scan()
        {
            List<string> siteCollectionsUrl = new List<string>();
            try
            {
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Starting to analyze on-premise environment");
                
                // JG CreateDataTableColumns(dt);
                Console.WriteLine(System.Environment.NewLine);
                Console.WriteLine("Starting to analyze on-premise environment");

                if (Scope == AnalysisScope.Farm)
                {
                    siteCollectionsUrl = QueryFarm();
                }
                else if (Scope == AnalysisScope.WebApplication)
                {
                    siteCollectionsUrl = GetAllWebAppSites(Url);

                }
                else if (Scope == AnalysisScope.SiteCollection)
                {
                    siteCollectionsUrl.Add(Url);
                }
                else if (Scope == AnalysisScope.SiteCollectionsUrls)
                {
                    siteCollectionsUrl = GetSiteCollectionsFromFile(FilePath);
                }

                List<WorkflowScanResult> workflows = FindWorkflows(siteCollectionsUrl);
                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");
                Logging.GetInstance().WriteToLogFile(Logging.Info, "TOTAL WORKFLOWS DISCOVERED : " + workflows.Count.ToString());
                Logging.GetInstance().WriteToLogFile(Logging.Info, "***********************************************************************");

                return workflows;

            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                return new List<WorkflowScanResult>();
            }
        }

        private List<string> GetSiteCollectionsFromFile(string filePath)
        {
            List<string> siteCollectionUrls = new List<string>();

            try
            {
                string line;
                System.IO.StreamReader file =
                    new System.IO.StreamReader(filePath);
                while ((line = file.ReadLine()) != null)
                {
                    //removes all extra spaces etc. 
                    siteCollectionUrls.Add(line.TrimEnd());
                }
                file.Close();
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }


            return siteCollectionUrls;

        }

        public List<string> GetAllWebAppSites(string url)
        {
            List<string> webAppSiteCollectionUrls = new List<string>();
            try
            {
                SPWebApplication objWebApp = null;
                objWebApp = SPWebApplication.Lookup(new Uri(url));
                if (objWebApp == null)
                {
                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    Console.WriteLine("Unable to obtain the object for the Web Application URL provided. Check to make sure the URL provided is correct.");
                    Console.ForegroundColor = ConsoleColor.White;
                    Logging.GetInstance().WriteToLogFile(Logging.Error, "Unable to obtain the object for the Web Application URL provided. SPWebApplication.Lookup(new Uri(Url)) returned NULL");
                }
                else
                {
                    foreach (SPSite site in objWebApp.Sites)
                    {
                        webAppSiteCollectionUrls.Add(site.Url);
                    }
                }


            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return webAppSiteCollectionUrls;
        }

        public List<string> QueryFarm()
        {
            List<string> farmSiteCollectionUrls = new List<string>();

            try
            {
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Starting to query the farm..");
                SPServiceCollection services = SPFarm.Local.Services;
                foreach (SPService curService in services)
                {
                    if (curService is SPWebService)
                    {
                        var webService = (SPWebService)curService;
                        if (curService.TypeName.Equals("Microsoft SharePoint Foundation Web Application"))
                        {
                            webService = (SPWebService)curService;
                            SPWebApplicationCollection webApplications = webService.WebApplications;
                            foreach (SPWebApplication webApplication in webApplications)
                            {
                                if (webApplication != null)
                                {
                                    if (false)
                                    {

                                    }
                                    else
                                    {
                                        foreach (SPSite site in webApplication.Sites)
                                        {
                                            try
                                            {
                                                farmSiteCollectionUrls.Add(site.Url);
                                            }
                                            catch (Exception ex)
                                            {
                                                Logging.GetInstance().WriteToLogFile(Logging.Error, "Errored! See log for details");
                                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                                                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }

            return farmSiteCollectionUrls;
        }

        public List<WorkflowScanResult> FindWorkflows(List<string> sitecollectionUrls)
        {
            List<WorkflowScanResult> results = new List<WorkflowScanResult>();

            try
            {
                foreach (string url in sitecollectionUrls)
                {
                    ClientContext siteClientContext = null;
                    if (Credential != null)
                    {
                        siteClientContext = CreateClientContext(url, Credential);
                    }
                    else
                    {
                        siteClientContext = CreateClientContext(url);
                    }
                    using (siteClientContext)
                    {
                        bool hasPermissions = false;

                        try

                        {
                            Console.WriteLine(string.Format("Processing: " + url));
                            siteClientContext.ExecuteQuery();
                            hasPermissions = true;
                        }
                        catch (WebException webException)
                        {
                            Console.WriteLine(string.Format(webException.Message.ToString() + " on " + url));
                            Logging.GetInstance().WriteToLogFile(Logging.Error, webException.Message.ToString() + " on " + url);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, webException.Message);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, webException.StackTrace);
                        }
                        catch (Microsoft.SharePoint.Client.ClientRequestException clientException)
                        {
                            Console.WriteLine(string.Format(clientException.Message.ToString() + " on " + url));
                            Logging.GetInstance().WriteToLogFile(Logging.Error, clientException.Message.ToString() + " on " + url);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, clientException.Message);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, clientException.StackTrace);
                        }
                        catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException unauthorizedException)
                        {
                            Console.WriteLine(string.Format(unauthorizedException.Message.ToString() + " on " + url));
                            Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message.ToString() + " on " + url);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message);
                            Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.StackTrace);
                        }

                        // Skip if did not have permissions
                        if (!hasPermissions)
                            continue;

                        Console.WriteLine(string.Format("Attempting to fetch all the sites and sub sites of  " + url));
                        IEnumerable<string> expandedSites = siteClientContext.Site.GetAllSubSites();

                        foreach (string site in expandedSites)
                        {
                            using (ClientContext ccWeb = siteClientContext.Clone(site))
                            {
                                try
                                {
                                    results = results.Union(FindWorkflowPerSite(ccWeb)).ToList();
                                }
                                catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException unauthorizedException)
                                {
                                    Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message);
                                    Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.StackTrace);
                                    Logging.GetInstance().WriteToLogFile(Logging.Error, unauthorizedException.Message.ToString() + " on " + url);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }

            return results;
        }

        public List<WorkflowScanResult> FindWorkflowPerSite(ClientContext cc)
        {
            try
            {
                var site = cc.Site;
                cc.Load(site);
                cc.ExecuteQuery();

                var web = cc.Web;
                cc.Load(web);
                cc.ExecuteQuery();

                WorkflowAnalyzer.Instance.LoadWorkflowDefaultActions();

                // Set up new WF discovery object for this site
                WorkflowDiscovery wfDisc = new WorkflowDiscovery();
                return wfDisc.DiscoverWorkflows(cc);
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                return new List<WorkflowScanResult>();
            }
        }

        internal ClientContext CreateClientContext(string url, NetworkCredential credential = null)
        {
            ClientContext cc = new ClientContext(url);
            try
            {
                if (credential == null)
                {
                    cc.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                }
                else
                {
                    cc.Credentials = credential;
                }
                Web web = cc.Web;
                cc.Load(web, website => website.Title);
                cc.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

                // Return ClientContext with url even though it will be unusable
                return new ClientContext(url);
            }
            return cc;
        }
    }
}

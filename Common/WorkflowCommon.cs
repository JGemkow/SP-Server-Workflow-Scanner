//using Microsoft.Deployment.Compression.Cab;
//using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Management.Automation;
using Microsoft.Online.SharePoint.TenantAdministration;
using Common.Models;

namespace Common
{
    
    public class Operations

    {
        /// <summary>
        /// SPO Properties
        /// </summary>
        public string DownloadPath { get; set; }
        public bool DownloadForms { get; set; }
        public string summaryFile = @"\WorkflowDiscovery.csv";
        public string logFolder = @"\Logs";
        public string downloadedFormsFolder = @"\DownloadedWorkflows";
        public string analysisFolder = @"\Analysis";
        public string summaryFolder = @"\Summary";
        public string analysisOutputFile = @"\WorkflowComparisonDetails.csv";
        public string compOutputFile = @"\WorkflowComparison.csv";
        public DataTable dt = new DataTable();


        public void SaveXamlFile(string xamlContent, Web web, string wfName, string scope, string folderPath)
        {
            try
            {
                string fileName = web.Id + "-" + wfName + "-" + scope+".xoml";
                string filePath = folderPath + "\\"+ fileName;
                System.IO.File.WriteAllText(filePath, xamlContent);
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }        
        }
        /// <summary>
        /// Create Data Table
        /// </summary>
        /// <param name="dt"></param>
        public void CreateDataTableColumns(DataTable dt)
        {
            try
            {
                dt.Columns.Add("SiteColID");
                dt.Columns.Add("SiteURL");
                dt.Columns.Add("ListTitle");
                dt.Columns.Add("ListUrl");
                dt.Columns.Add("ContentTypeId");
                dt.Columns.Add("ContentTypeName");
                dt.Columns.Add("Scope");
                dt.Columns.Add("Version");
                dt.Columns.Add("WFTemplateName");
                dt.Columns.Add("WorkFlowName");
                dt.Columns.Add("IsOOBWorkflow");
                dt.Columns.Add("WFID");
                dt.Columns.Add("WebID");
                dt.Columns.Add("WebURL");
                dt.Columns.Add("Enabled");
                dt.Columns.Add("HasSubscriptions");
                dt.Columns.Add("ConsiderUpgradingToFlow");
                dt.Columns.Add("ToFLowMappingPercentage");
                dt.Columns.Add("UsedActions");
                dt.Columns.Add("ActionCount");
                dt.Columns.Add("AllowManual");
                dt.Columns.Add("AutoStartChange");
                dt.Columns.Add("AutoStartCreate");
                dt.Columns.Add("LastDefinitionModifiedDate");
                dt.Columns.Add("LastSubsrciptionModifiedDate");
                dt.Columns.Add("AssociationData");
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
        }

        /// <summary>
        /// Create Data Table
        /// </summary>
        /// <param name="dt"></param>
        public void AddRowToDataTable(WorkflowScanResult workflowScanResult, DataTable dt, string version, string scope, string wfName, string wfID, bool IsOOBWF, Web web)
        {
            DataRow dr = dt.NewRow();
            try
            {
                if (workflowScanResult.SiteColUrl == null && web.Url != null)
                {
                    dr["SiteURL"] = web.Url;
                    dr["SiteColID"] = web.Id;
                }
                else
                dr["SiteURL"] = workflowScanResult.SiteColUrl;
                dr["WebURL"] = web.Url;
                dr["ListTitle"] = workflowScanResult.ListTitle;
                dr["ListUrl"] = workflowScanResult.ListUrl;
                dr["ContentTypeId"] = workflowScanResult.ContentTypeId;
                dr["ContentTypeName"] = workflowScanResult.ContentTypeName;
                dr["Scope"] = scope;
                dr["Version"] = version;
                dr["WFTemplateName"] = wfName;
                dr["WorkFlowName"] = workflowScanResult.SubscriptionName;
                dr["IsOOBWorkflow"] = IsOOBWF;
                dr["Enabled"] = workflowScanResult.Enabled;   // adding for is enabled 
                dr["WFID"] = wfID;
                dr["WebID"] = web.Id;
                dr["HasSubscriptions"] = workflowScanResult.HasSubscriptions;   // adding for subscriptions 
                string sUsedActions = "";
                // AM need to refactor into a helper function
                if (workflowScanResult.UsedActions != null)
                {
                    foreach (var item in workflowScanResult.UsedActions)
                    {
                        sUsedActions = item.ToString()+";"+ sUsedActions;
                    }
                }
                dr["ToFLowMappingPercentage"] = workflowScanResult.ToFLowMappingPercentage;   // adding for percentange upgradable to flow 
                dr["ConsiderUpgradingToFlow"] = workflowScanResult.ConsiderUpgradingToFlow;   // adding for consider upgrading to flow 
                dr["UsedActions"] = sUsedActions;   // adding for UsedActions
                dr["ActionCount"] = workflowScanResult.ActionCount;   // adding for ActionCount
                dr["AllowManual"] = workflowScanResult.AllowManual;
                dr["AutoStartChange"] = workflowScanResult.AutoStartChange;
                dr["AutoStartCreate"] = workflowScanResult.AutoStartCreate;
                dr["LastDefinitionModifiedDate"] = workflowScanResult.LastDefinitionEdit;
                dr["LastSubsrciptionModifiedDate"] = workflowScanResult.LastSubscriptionEdit;
                dr["AssociationData"] = workflowScanResult.AssociationData;

                dt.Rows.Add(dr);
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
        }

        /// <summary>
        /// Creates 3 levels of folder at the downloaded path provided by the user
        /// 1. Analysis
        /// 2. DownloadedForms
        /// 3. Summary
        /// </summary>
        /// <param name="downloadPath"></param>
        public void CreateDirectoryStructure(string downloadPath)
        {
            try
            {
                if (!Directory.Exists(string.Concat(downloadPath, analysisFolder)))
                {
                    DirectoryInfo analysisFolder1 = System.IO.Directory.CreateDirectory(string.Concat(downloadPath, analysisFolder));
                    Logging.GetInstance().WriteToLogFile(Logging.Info, "Analysis folder created");
                }
                if (!Directory.Exists(string.Concat(downloadPath, downloadedFormsFolder)))
                {
                    DirectoryInfo downloadedFormsFolder1 = System.IO.Directory.CreateDirectory(string.Concat(downloadPath, downloadedFormsFolder));
                    Logging.GetInstance().WriteToLogFile(Logging.Info, "DownloadedForms folder created");

                }
                if (!Directory.Exists(string.Concat(downloadPath, summaryFolder)))
                {
                    DirectoryInfo summaryFolder1 = System.IO.Directory.CreateDirectory(string.Concat(downloadPath, summaryFolder));
                    Logging.GetInstance().WriteToLogFile(Logging.Info, "Summary folder created");

                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
        }

        public DataTable ConvertCSVToDataTable(string csvLocation)
        {
            DataTable dt = new DataTable();
            try
            {
                var Lines = System.IO.File.ReadAllLines(csvLocation);
                string[] Fields;
                Fields = Lines[0].Split(new char[] { ',' });
                int Cols = Fields.GetLength(0);
                for (int i = 0; i < Cols; i++)
                    dt.Columns.Add(Fields[i].ToLower(), typeof(string));
                DataRow Row;
                for (int i = 1; i < Lines.GetLength(0); i++)
                {
                    Fields = Lines[i].Split(new char[] { ',' });
                    Row = dt.NewRow();
                    for (int f = 0; f < Cols; f++)
                        Row[f] = Fields[f].Split('"')[1];
                    dt.Rows.Add(Row);
                }
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
            return dt;
        }


        /// <summary>
        /// The DataSet returned from the content database is stored in a datatable
        /// The datatable is then saved into a CSV file that gets stored in the Summary folder
        /// that gets created at the location of download path supplied by the users
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="filePath"></param>
        public void WriteToCsvFile(DataTable dataTable, string filePath)
        {
            try
            {
                Logging.GetInstance().WriteToLogFile(Logging.Info, "Preparing to create the CSV at " + filePath);
                StringBuilder fileContent = new StringBuilder();

                foreach (var col in dataTable.Columns)
                {
                    fileContent.Append(col.ToString() + ",");
                }

                fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);

                foreach (DataRow dr in dataTable.Rows)
                {
                    foreach (var column in dr.ItemArray)
                    {
                        fileContent.Append("\"" + column.ToString() + "\",");
                    }

                    fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);
                }

                System.IO.File.WriteAllText(filePath, fileContent.ToString());
                Logging.GetInstance().WriteToLogFile(Logging.Info, string.Format("CSV File created at {0}", filePath));
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="s"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void WriteProgress(string s, int x, int y)
        {
            int origRow = Console.CursorTop;
            int origCol = Console.CursorLeft;
            // Console.WindowWidth = 10;  // this works. 
            int width = Console.WindowWidth;
            //x = x % width;
            try
            {
                Console.SetCursorPosition(x, y);
                //Console.SetCursorPosition(origCol, origRow);
                Console.Write(s);
            }
            catch (ArgumentOutOfRangeException e)
            {

            }
            finally
            {
                try
                {
                    Console.SetCursorPosition(origRow, origCol);
                }
                catch (ArgumentOutOfRangeException e)
                {
                }
            }
        }

        internal ClientContext CreateClientContext(string url, string username, SecureString password)
        {
            try
            {
                var credentials = new SharePointOnlineCredentials(
                                       username,
                                       password);

                return new ClientContext(url)
                {
                    Credentials = credentials
                };
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                return new ClientContext(url)
                {
                };
            }


        }


        public List<string> GetAllTenantSites(string TenantName, PSCredential Credential)
        {
            List<string> sites = new List<string>();
            try
            {
                string tenantAdminUrl = "https://" + TenantName + "-admin.sharepoint.com/";
                ClientContext ctx = null;
                ctx = CreateClientContext(tenantAdminUrl, Credential.UserName, Credential.Password);
                Tenant tenant = new Tenant(ctx);
                SPOSitePropertiesEnumerable siteProps = tenant.GetSitePropertiesFromSharePoint("0", true);
                ctx.Load(siteProps);
                ctx.ExecuteQuery();
                int count = 0;
                foreach (var site in siteProps)
                {
                    sites.Add(site.Url);
                    count++;
                }
                Console.WriteLine("Total Site {0}", count);
                return sites;
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

            }
            return sites;
        }

        /// <summary>
        /// Open the site collection file and store in a collection variable
        /// Read the file and display it line by line.  
        /// </summary>
        /// <param name="sitecollectionUrls"></param>
        public void ReadInfoPathOnlineSiteCollection(List<string> sitecollectionUrls, string filePath)
        {
            try
            {
                int counter = 0;
                string line;
                System.IO.StreamReader file =
                    new System.IO.StreamReader(filePath);
                while ((line = file.ReadLine()) != null)
                {
                    //removes all extra spaces etc. 
                    sitecollectionUrls.Add(line.TrimEnd());
                    counter++;
                }
                file.Close();
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }
        }
    }
}


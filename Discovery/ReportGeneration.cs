﻿using System;
using System.IO;
using System.Text;
using System.Data;
using Microsoft.SharePoint.Client;
using WorkflowScanner.Models;
using System.Linq;
using System.Collections.Generic;
using Microsoft.IdentityModel.Protocols.WSIdentity;
using PnP.Core.Model.SharePoint;
using System.Xml.Linq;

namespace Common
{
    public static class WorkflowScanResultExtensions
    {
        public static DataTable ToDataTable(this List<WorkflowScanResult> results)
        {
            DataTable dt = new DataTable();

            foreach (WorkflowScanResult workflowScanResult in results)
            {
                DataRow dr = dt.NewRow();
                try
                {
                    dr["SiteURL"] = workflowScanResult.SiteURL;
                    dr["SiteCollID"] = workflowScanResult.SiteCollID;
                    dr["WebURL"] = workflowScanResult.WebUrl;
                    dr["ListTitle"] = workflowScanResult.ListTitle;
                    dr["ListUrl"] = workflowScanResult.ListUrl;
                    dr["ContentTypeId"] = workflowScanResult.ContentTypeId;
                    dr["ContentTypeName"] = workflowScanResult.ContentTypeName;
                    dr["Scope"] = workflowScanResult.Scope;
                    dr["Version"] = workflowScanResult.Version;
                    dr["WFTemplateName"] = workflowScanResult.WorkflowTemplateName;
                    dr["WorkFlowName"] = workflowScanResult.SubscriptionName;
                    dr["IsOOBWorkflow"] = workflowScanResult.IsOOBWorkflow;
                    dr["Enabled"] = workflowScanResult.Enabled;   // adding for is enabled 
                    dr["WFID"] = workflowScanResult.DefinitionId;
                    dr["WebID"] = workflowScanResult.WebId;
                    dr["HasSubscriptions"] = workflowScanResult.HasSubscriptions;   // adding for subscriptions 
                    dr["UsedActions"] = FormatUsedActions(workflowScanResult.UsedActions);   // adding for UsedActions
                    dr["ToFLowMappingPercentage"] = workflowScanResult.ToFLowMappingPercentage;   // adding for percentange upgradable to flow 
                    dr["ConsiderUpgradingToFlow"] = workflowScanResult.ConsiderUpgradingToFlow;   // adding for consider upgrading to flow 
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

            return dt;

        }

        private static string FormatUsedActions(List<string> actions)
        {
            return actions != null ? actions.Aggregate((current, next) => current + "; " + next) : "";
        }
    }


    public class ReportGeneration

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


        /// <summary>
        /// Create Data Table
        /// </summary>
        /// <param name="dt"></param>
        public void AddRowToDataTable(WorkflowScanResult workflowScanResult, DataTable dt, string version, string scope, string wfName, string wfID, bool IsOOBWF, Web web)
        {
            DataRow dr = dt.NewRow();
            try
            {
                if (workflowScanResult.SiteCollID == null && web.Url != null)
                {
                    dr["SiteURL"] = web.Url;
                    dr["SiteColID"] = web.Id;
                }
                else
                dr["SiteURL"] = workflowScanResult.SiteCollID;
                dr["WebURL"] = web.Url;
                dr["ListTitle"] = workflowScanResult.ListTitle;
                dr["ListUrl"] = workflowScanResult.ListUrl;
                dr["ContentTypeId"] = workflowScanResult.ContentTypeId;
                dr["ContentTypeName"] = workflowScanResult.ContentTypeName;
                dr["Scope"] = workflowScanResult.Scope;
                dr["Version"] = version;
                dr["WFTemplateName"] = wfName;
                dr["WorkFlowName"] = workflowScanResult.SubscriptionName;
                dr["IsOOBWorkflow"] = IsOOBWF;
                dr["Enabled"] = workflowScanResult.Enabled;   // adding for is enabled 
                dr["WFID"] = workflowScanResult.WFID;
                dr["WebID"] = web.Id;
                dr["HasSubscriptions"] = workflowScanResult.HasSubscriptions;   // adding for subscriptions 
                dr["UsedActions"] = FormatUsedActions(workflowScanResult.UsedActions);   // adding for UsedActions
                dr["ToFLowMappingPercentage"] = workflowScanResult.ToFLowMappingPercentage;   // adding for percentange upgradable to flow 
                dr["ConsiderUpgradingToFlow"] = workflowScanResult.ConsiderUpgradingToFlow;   // adding for consider upgrading to flow 
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

        public string FormatUsedActions(List<string> actions)
        {
            return actions != null ? actions.Aggregate((current, next) => current + "; " + next) : "";
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
    }
}


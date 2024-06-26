﻿using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using Common;
using System.Data;
using Common.Models;
using WorkflowScanner.Models;
using Microsoft.SharePoint.Workflow;

namespace Discovery
{
    public class WorkflowDiscovery
    {

        private class SP2010WorkFlowAssociation
        {
            public string Scope { get; set; }
            public WorkflowAssociation WorkflowAssociation { get; set; }
            public List AssociatedList { get; set; }
            public ContentType AssociatedContentType { get; set; }
        }

        private static readonly string[] OOBWorkflowIDStarts = new string[]
        {
            "e43856d2-1bb4-40ef-b08b-016d89a00",    // Publishing approval
            "3bfb07cb-5c6a-4266-849b-8d6711700409", // Collect feedback - 2010
            "46c389a4-6e18-476c-aa17-289b0c79fb8f", // Collect feedback
            "77c71f43-f403-484b-bcb2-303710e00409", // Collect signatures - 2010
            "2f213931-3b93-4f81-b021-3022434a3114", // Collect signatures
            "8ad4d8f0-93a7-4941-9657-cf3706f00409", // Approval - 2010
            "b4154df4-cc53-4c4f-adef-1ecf0b7417f6", // Translation management
            "c6964bff-bf8d-41ac-ad5e-b61ec111731a", // Three state
            "c6964bff-bf8d-41ac-ad5e-b61ec111731c", // Approval
            "dd19a800-37c1-43c0-816d-f8eb5f4a4145", // Disposition approval
        };

        List<SP2010WorkFlowAssociation> sp2010WorkflowAssociations = new List<SP2010WorkFlowAssociation>(20);
        private List workflowList;

        public List<WorkflowScanResult>DiscoverWorkflows(ClientContext cc)
        {
            List<WorkflowScanResult> results = new List<WorkflowScanResult>();
            try
            {
                Web web = cc.Web;

                // Pre-load needed properties in a single call
                cc.Load(web, w => w.Id, w => w.Language, w => w.ServerRelativeUrl, w => w.Url, w => w.WorkflowTemplates, w => w.WorkflowAssociations);
                cc.Load(web, p => p.ContentTypes.Include(ct => ct.WorkflowAssociations, ct => ct.Name, ct => ct.StringId));
                cc.Load(web, p => p.Lists.Include(li => li.Id, li => li.Title, li => li.Hidden, li => li.DefaultViewUrl, li => li.BaseTemplate, li => li.RootFolder.ServerRelativeUrl, li => li.ItemCount, li => li.WorkflowAssociations, li => li.ContentTypesEnabled, li => li.ContentTypes.Include(ct => ct.WorkflowAssociations, ct => ct.Name, ct => ct.StringId)));
                cc.Load(cc.Site, p => p.RootWeb);
                cc.Load(cc.Site.RootWeb, p => p.Lists.Include(li => li.Id, li => li.Title, li => li.Hidden, li => li.DefaultViewUrl, li => li.BaseTemplate, li => li.RootFolder.ServerRelativeUrl, li => li.ItemCount, li => li.WorkflowAssociations, li => li.ContentTypesEnabled, li => li.ContentTypes.Include(ct => ct.WorkflowAssociations, ct => ct.Name, ct => ct.StringId)));
                
                cc.ExecuteQuery();
                var lists = web.Lists;

                #region Site, reusable and list level 2013 workflow
                // *******************************************
                // Site, reusable and list level 2013 workflow
                // *******************************************

                // Retrieve the 2013 site level workflow definitions (including unpublished ones)
                WorkflowDefinition[] siteWFDefinitions = null;
                // Retrieve the 2013 site level workflow subscriptions
                WorkflowSubscription[] siteWFSubscriptions = null;
                ReportGeneration reportGen = new ReportGeneration();
                try
                {
                    var servicesManager = new WorkflowServicesManager(web.Context, web);
                    var deploymentService = servicesManager.GetWorkflowDeploymentService();
                    var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

                    var definitions = deploymentService.EnumerateDefinitions(false);
                    web.Context.Load(definitions);
                    var subscriptions = subscriptionService.EnumerateSubscriptions();
                    web.Context.Load(subscriptions);

                    web.Context.ExecuteQuery();

                    siteWFDefinitions = definitions.ToArray();
                    siteWFSubscriptions = subscriptions.ToArray();
                }
                catch (ServerException ex)
                {
                    // If there is no workflow service present in the farm this method will throw an error. 
                    Logging.GetInstance().WriteToLogFile(Logging.Info, "No workflow service present in farm for 2013 WFM.");
                    Logging.GetInstance().WriteToLogFile(Logging.Info, ex.Message);
                    Logging.GetInstance().WriteToLogFile(Logging.Info, ex.StackTrace);
                }
                #endregion

                #region If SP2013 site scoped workflows are found
                // We've found SP2013 site scoped workflows
                if (siteWFDefinitions != null && siteWFDefinitions.Count() > 0)
                {
                    foreach (var siteWFDefinition in siteWFDefinitions.Where(p => p.RestrictToType != null && (p.RestrictToType.Equals("site", StringComparison.InvariantCultureIgnoreCase) || p.RestrictToType.Equals("universal", StringComparison.InvariantCultureIgnoreCase))))
                    {
                        // Check if this workflow is also in use
                        var siteWorkflowSubscriptions = siteWFSubscriptions.Where(p => p.DefinitionId.Equals(siteWFDefinition.Id));

                        // Perform workflow analysis
                        var workFlowAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowDefinition(siteWFDefinition.Xaml, WorkflowTypes.SP2013);
                        
                        var workFlowTriggerAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowTriggers(GetWorkflowPropertyBool(siteWFDefinition.Properties, "SPDConfig.StartOnCreate"), GetWorkflowPropertyBool(siteWFDefinition.Properties, "SPDConfig.StartOnChange"), GetWorkflowPropertyBool(siteWFDefinition.Properties, "SPDConfig.StartManually"));

                        if (siteWorkflowSubscriptions.Count() > 0)
                        {
                            foreach (var siteWorkflowSubscription in siteWorkflowSubscriptions)
                            {
                                results.Add(new WorkflowScanResult()
                                {
                                    ListTitle = "",
                                    ListUrl = "",
                                    ContentTypeId = "",
                                    ContentTypeName = "",
                                    Version = "2013",
                                    Scope = "Site",
                                    RestrictToType = siteWFDefinition.RestrictToType,
                                    DefinitionName = siteWFDefinition.DisplayName,
                                    DefinitionDescription = siteWFDefinition.Description,
                                    SubscriptionName = siteWorkflowSubscription.Name,
                                    HasSubscriptions = true,
                                    Enabled = siteWorkflowSubscription.Enabled,
                                    DefinitionId = siteWFDefinition.Id,
                                    IsOOBWorkflow = false,
                                    SubscriptionId = siteWorkflowSubscription.Id,
                                    UsedActions = workFlowAnalysisResult?.WorkflowActions,
                                    ActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.ActionCount : 0,
                                    UsedTriggers = workFlowTriggerAnalysisResult?.WorkflowTriggers,
                                    UnsupportedActionsInFlow = workFlowAnalysisResult?.UnsupportedActions,
                                    CreatedBy = siteWFDefinition.Properties["vti_author"],
                                    ModifiedBy = siteWFDefinition.Properties["ModifiedBy"],
                                    UnsupportedActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.UnsupportedAccountCount : 0,
                                    LastDefinitionEdit = GetWorkflowPropertyDateTime(siteWFDefinition.Properties, "Definition.ModifiedDateUTC"),
                                    LastSubscriptionEdit = GetWorkflowPropertyDateTime(siteWorkflowSubscription.PropertyDefinitions, "SharePointWorkflowContext.Subscription.ModifiedDateUTC"),

                                    WorkflowTemplateName = siteWFDefinition.DisplayName,
                                    WFID = siteWFDefinition.Id.ToString(),
                                    WebId = web.Id.ToString(),
                                    WebUrl = web.Url,
                                    SiteURL = cc.Site.Url,
                                    SiteCollID = cc.Site.Id.ToString()
                                   
                                });
                            }
                        }
                        else
                        {
                            results.Add(new WorkflowScanResult()
                            {
                                ListTitle = "",
                                ListUrl = "",
                                ContentTypeId = "",
                                ContentTypeName = "",
                                Version = "2013",
                                Scope = "Site",
                                RestrictToType = siteWFDefinition.RestrictToType,
                                DefinitionName = siteWFDefinition.DisplayName,
                                DefinitionDescription = siteWFDefinition.Description,
                                SubscriptionName = "",
                                HasSubscriptions = false,
                                Enabled = false,
                                DefinitionId = siteWFDefinition.Id,
                                IsOOBWorkflow = false,
                                SubscriptionId = Guid.Empty,
                                UsedActions = workFlowAnalysisResult?.WorkflowActions,
                                ActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.ActionCount : 0,
                                UnsupportedActionsInFlow = workFlowAnalysisResult?.UnsupportedActions,
                                UnsupportedActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.UnsupportedAccountCount : 0,
                                UsedTriggers = workFlowTriggerAnalysisResult?.WorkflowTriggers,
                                LastDefinitionEdit = GetWorkflowPropertyDateTime(siteWFDefinition.Properties, "Definition.ModifiedDateUTC"),
                                WorkflowTemplateName = siteWFDefinition.DisplayName,
                                WFID = siteWFDefinition.Id.ToString(),
                                WebId = web.Id.ToString(),
                                WebUrl = web.Url,
                                SiteURL = cc.Site.Url,
                                SiteCollID = cc.Site.Id.ToString()

                            });
                        }
                    }
                }
                #endregion

                #region If SP2013 list scoped workflows are found
                // We've found SP2013 list scoped workflows
                if (siteWFDefinitions != null && siteWFDefinitions.Count() > 0)
                {
                    foreach (var listWFDefinition in siteWFDefinitions.Where(p => p.RestrictToType != null && (p.RestrictToType.Equals("list", StringComparison.InvariantCultureIgnoreCase) || p.RestrictToType.Equals("universal", StringComparison.InvariantCultureIgnoreCase))))
                    {
                        // Check if this workflow is also in use
                        var listWorkflowSubscriptions = siteWFSubscriptions.Where(p => p.DefinitionId.Equals(listWFDefinition.Id));

                        // Perform workflow analysis
                        var workFlowAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowDefinition(listWFDefinition.Xaml, WorkflowTypes.SP2013);
                        
                        var workFlowTriggerAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowTriggers(GetWorkflowPropertyBool(listWFDefinition.Properties, "SPDConfig.StartOnCreate"), GetWorkflowPropertyBool(listWFDefinition.Properties, "SPDConfig.StartOnChange"), GetWorkflowPropertyBool(listWFDefinition.Properties, "SPDConfig.StartManually"));

                        if (listWorkflowSubscriptions.Count() > 0)
                        {
                            foreach (var listWorkflowSubscription in listWorkflowSubscriptions)
                            {
                                Guid associatedListId = Guid.Empty;
                                string associatedListTitle = "";
                                string associatedListUrl = "";
                                if (Guid.TryParse(GetWorkflowProperty(listWorkflowSubscription, "Microsoft.SharePoint.ActivationProperties.ListId"), out Guid associatedListIdValue))
                                {
                                    associatedListId = associatedListIdValue;

                                    // Lookup this list and update title and url
                                    var listLookup = lists.Where(p => p.Id.Equals(associatedListId)).FirstOrDefault();
                                    if (listLookup != null)
                                    {
                                        associatedListTitle = listLookup.Title;
                                        associatedListUrl = listLookup.RootFolder.ServerRelativeUrl;
                                    }
                                }

                                results.Add(new WorkflowScanResult()
                                {
                                    ListTitle = associatedListTitle,
                                    ListUrl = associatedListUrl,
                                    ListId = associatedListId,
                                    ContentTypeId = "",
                                    ContentTypeName = "",
                                    Version = "2013",
                                    Scope = "List",
                                    RestrictToType = listWFDefinition.RestrictToType,
                                    DefinitionName = listWFDefinition.DisplayName,
                                    DefinitionDescription = listWFDefinition.Description,
                                    SubscriptionName = listWorkflowSubscription.Name,
                                    HasSubscriptions = true,
                                    Enabled = listWorkflowSubscription.Enabled,
                                    DefinitionId = listWFDefinition.Id,
                                    IsOOBWorkflow = false,
                                    SubscriptionId = listWorkflowSubscription.Id,
                                    UsedActions = workFlowAnalysisResult?.WorkflowActions,
                                    ActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.ActionCount : 0,
                                    UsedTriggers = workFlowTriggerAnalysisResult?.WorkflowTriggers,
                                    UnsupportedActionsInFlow = workFlowAnalysisResult?.UnsupportedActions,
                                    CreatedBy = listWFDefinition.Properties["vti_author"],
                                    ModifiedBy = listWFDefinition.Properties["ModifiedBy"],
                                    UnsupportedActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.UnsupportedAccountCount : 0,
                                    LastDefinitionEdit = GetWorkflowPropertyDateTime(listWFDefinition.Properties, "Definition.ModifiedDateUTC"),
                                    LastSubscriptionEdit = GetWorkflowPropertyDateTime(listWorkflowSubscription.PropertyDefinitions, "SharePointWorkflowContext.Subscription.ModifiedDateUTC"),
                                    WorkflowTemplateName = listWFDefinition.DisplayName,
                                    WFID = listWFDefinition.Id.ToString(),
                                    WebId = web.Id.ToString(),
                                    WebUrl = web.Url,
                                    SiteURL = cc.Site.Url,
                                    SiteCollID = cc.Site.Id.ToString()
                                });
                            }
                        }
                        else
                        {
                            results.Add(new WorkflowScanResult()
                            {
                                ListTitle = "",
                                ListUrl = "",
                                ListId = Guid.Empty,
                                ContentTypeId = "",
                                ContentTypeName = "",
                                Version = "2013",
                                Scope = "List",
                                RestrictToType = listWFDefinition.RestrictToType,
                                DefinitionName = listWFDefinition.DisplayName,
                                DefinitionDescription = listWFDefinition.Description,
                                SubscriptionName = "",
                                HasSubscriptions = false,
                                Enabled = false,
                                DefinitionId = listWFDefinition.Id,
                                IsOOBWorkflow = false,
                                SubscriptionId = Guid.Empty,
                                UsedActions = workFlowAnalysisResult?.WorkflowActions,
                                ActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.ActionCount : 0,
                                UsedTriggers = workFlowTriggerAnalysisResult?.WorkflowTriggers,
                                UnsupportedActionsInFlow = workFlowAnalysisResult?.UnsupportedActions,
                                UnsupportedActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.UnsupportedAccountCount : 0,
                                LastDefinitionEdit = GetWorkflowPropertyDateTime(listWFDefinition.Properties, "Definition.ModifiedDateUTC"),
                                WorkflowTemplateName = listWFDefinition.DisplayName,
                                WFID = listWFDefinition.Id.ToString(),
                                WebId = web.Id.ToString(),
                                WebUrl = web.Url,
                                SiteURL = cc.Site.Url,
                                SiteCollID = cc.Site.Id.ToString()
                            });

                        }
                    }
                }
                #endregion

                #region Find all Site, list and content type level 2010 workflows
                // ***********************************************
                // Site, list and content type level 2010 workflow
                // ***********************************************

                // Find all places where we have workflows associated (=subscribed) to SharePoint objects
                if (web.WorkflowAssociations.Count > 0)
                {
                    foreach (var workflowAssociation in web.WorkflowAssociations)
                    {
                        this.sp2010WorkflowAssociations.Add(new SP2010WorkFlowAssociation() { Scope = "Site", WorkflowAssociation = workflowAssociation });
                    }
                }

                foreach (var list in lists.Where(p => p.WorkflowAssociations.Count > 0))
                {
                    foreach (var workflowAssociation in list.WorkflowAssociations)
                    {
                        this.sp2010WorkflowAssociations.Add(new SP2010WorkFlowAssociation() { Scope = "List", WorkflowAssociation = workflowAssociation, AssociatedList = list });
                    }
                }

                foreach (var list in lists.Where(p => p.ContentTypesEnabled))
                {
                    foreach (var listContentType in list.ContentTypes.Where(p => p.WorkflowAssociations.Count > 0))
                    {
                        foreach (var workflowAssociation in listContentType.WorkflowAssociations)
                        {
                            this.sp2010WorkflowAssociations.Add(new SP2010WorkFlowAssociation() { Scope = "ContentType", WorkflowAssociation = workflowAssociation, AssociatedContentType = listContentType, AssociatedList = list });
                        }
                    }
                }
                foreach (var ct in web.ContentTypes.Where(p => p.WorkflowAssociations.Count > 0))
                {
                    foreach (var workflowAssociation in ct.WorkflowAssociations)
                    {
                        this.sp2010WorkflowAssociations.Add(new SP2010WorkFlowAssociation() { Scope = "ContentType", WorkflowAssociation = workflowAssociation, AssociatedContentType = ct });
                    }
                }
                #endregion
                
                #region Process 2010 Workflows
                // Process 2010 worflows                
                List<Guid> processedWorkflowAssociations = new List<Guid>(this.sp2010WorkflowAssociations.Count);              

                if (web.WorkflowTemplates.Count > 0)
                {

                    // Process the templates
                    foreach (var workflowTemplate in web.WorkflowTemplates)
                    {
                        // do we have workflows associated for this template?
                        var associatedWorkflows = this.sp2010WorkflowAssociations.Where(p => p.WorkflowAssociation.BaseId.Equals(workflowTemplate.Id));
                        if (associatedWorkflows.Count() > 0)
                        {
                            // Perform workflow analysis
                            // If returning null than this workflow template was an OOB workflow one
                            WorkflowActionAnalysis workFlowAnalysisResult = null;
                            var loadedWorkflow = LoadWorkflowDefinition(cc, workflowTemplate);
                            if (!string.IsNullOrEmpty(loadedWorkflow?.Item1))
                            {
                                workFlowAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowDefinition(loadedWorkflow.Item1, WorkflowTypes.SP2010);
                            }

                            foreach (var associatedWorkflow in associatedWorkflows)
                            {
                                processedWorkflowAssociations.Add(associatedWorkflow.WorkflowAssociation.Id);

                                // Skip previous versions of a workflow
                                // TODO: non-english sites will use another string
                                if (associatedWorkflow.WorkflowAssociation.Name.Contains("(Previous Version:"))
                                {
                                    continue;
                                }

                                var workFlowTriggerAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowTriggers(associatedWorkflow.WorkflowAssociation.AutoStartCreate, associatedWorkflow.WorkflowAssociation.AutoStartChange, associatedWorkflow.WorkflowAssociation.AllowManual);

                                results.Add(new WorkflowScanResult()
                                {
                                    ListTitle = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.Title : "",
                                    ListUrl = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.RootFolder.ServerRelativeUrl : "",
                                    ListId = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.Id : Guid.Empty,
                                    ContentTypeId = associatedWorkflow.AssociatedContentType != null ? associatedWorkflow.AssociatedContentType.StringId : "",
                                    ContentTypeName = associatedWorkflow.AssociatedContentType != null ? associatedWorkflow.AssociatedContentType.Name : "",
                                    Version = "2010",
                                    Scope = associatedWorkflow.Scope,
                                    RestrictToType = "N/A",
                                    DefinitionName = workflowTemplate.Name,
                                    DefinitionDescription = workflowTemplate.Description,
                                    SubscriptionName = associatedWorkflow.WorkflowAssociation.Name,
                                    HasSubscriptions = true,
                                    Enabled = associatedWorkflow.WorkflowAssociation.Enabled,
                                    DefinitionId = workflowTemplate.Id,
                                    IsOOBWorkflow = IsOOBWorkflow(workflowTemplate.Id.ToString()),
                                    SubscriptionId = associatedWorkflow.WorkflowAssociation.Id,
                                    UsedActions = workFlowAnalysisResult?.WorkflowActions,
                                    ActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.ActionCount : 0,
                                    UsedTriggers = workFlowTriggerAnalysisResult?.WorkflowTriggers,
                                    UnsupportedActionsInFlow = workFlowAnalysisResult?.UnsupportedActions,
                                    UnsupportedActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.UnsupportedAccountCount : 0,
                                    LastDefinitionEdit = loadedWorkflow != null ? loadedWorkflow.Item2 : associatedWorkflow.WorkflowAssociation.Modified,
                                    LastSubscriptionEdit = associatedWorkflow.WorkflowAssociation.Modified,
                                    AssociationData = associatedWorkflow.WorkflowAssociation.AssociationData,
                                    AllowManual = associatedWorkflow.WorkflowAssociation.AllowManual,
                                    AutoStartChange = associatedWorkflow.WorkflowAssociation.AutoStartChange,
                                    AutoStartCreate = associatedWorkflow.WorkflowAssociation.AutoStartCreate,
                                    WorkflowTemplateName = workflowTemplate.Name,
                                    WFID = workflowTemplate.Id.ToString(),
                                    WebId = web.Id.ToString(),
                                    WebUrl = web.Url,
                                    SiteURL = cc.Site.Url,
                                    SiteCollID = cc.Site.Id.ToString()
                                });
                            }
                        }
                        else
                        {
                            // Only add non OOB workflow templates when there's no associated workflow - makes the dataset smaller
                            if (!IsOOBWorkflow(workflowTemplate.Id.ToString()))
                            {
                                // Perform workflow analysis
                                WorkflowActionAnalysis workFlowAnalysisResult = null;
                                var loadedWorkflow = LoadWorkflowDefinition(cc, workflowTemplate);
                                if (!string.IsNullOrEmpty(loadedWorkflow?.Item1))
                                {
                                    workFlowAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowDefinition(loadedWorkflow.Item1, WorkflowTypes.SP2010);
                                }

                                var workFlowTriggerAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowTriggers(workflowTemplate.AutoStartCreate, workflowTemplate.AutoStartChange, workflowTemplate.AllowManual);

                                results.Add(new WorkflowScanResult()
                                {
                                    ListTitle = "",
                                    ListUrl = "",
                                    ListId = Guid.Empty,
                                    ContentTypeId = "",
                                    ContentTypeName = "",
                                    Version = "2010",
                                    Scope = "",
                                    RestrictToType = "N/A",
                                    DefinitionName = workflowTemplate.Name,
                                    DefinitionDescription = workflowTemplate.Description,
                                    SubscriptionName = "",
                                    HasSubscriptions = false,
                                    Enabled = false,
                                    DefinitionId = workflowTemplate.Id,
                                    IsOOBWorkflow = IsOOBWorkflow(workflowTemplate.Id.ToString()),
                                    SubscriptionId = Guid.Empty,
                                    UsedActions = workFlowAnalysisResult?.WorkflowActions,
                                    ActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.ActionCount : 0,
                                    UsedTriggers = workFlowTriggerAnalysisResult?.WorkflowTriggers,
                                    LastDefinitionEdit = loadedWorkflow != null ? loadedWorkflow.Item2 : DateTime.MinValue,
                                    AssociationData = "",
                                    WorkflowTemplateName = workflowTemplate.Name,
                                    WFID = workflowTemplate.Id.ToString(),
                                    WebId = web.Id.ToString(),
                                    WebUrl = web.Url,
                                    SiteURL = cc.Site.Url,
                                    SiteCollID = cc.Site.Id.ToString()
                                });
                            }
                        }
                    }
                }
                #endregion

                #region All other Workflows
                // Are there associated workflows for which we did not find a template (especially when the WF is created for a list)
                foreach (var associatedWorkflow in this.sp2010WorkflowAssociations)
                {
                    if (!processedWorkflowAssociations.Contains(associatedWorkflow.WorkflowAssociation.Id))
                    {
                        // Skip previous versions of a workflow
                        // TODO: non-english sites will use another string
                        if (associatedWorkflow.WorkflowAssociation.Name.Contains("(Previous Version:"))
                        {
                            continue;
                        }

                        // Perform workflow analysis
                        WorkflowActionAnalysis workFlowAnalysisResult = null;
                        var loadedWorkflow = LoadWorkflowDefinition(cc, associatedWorkflow.WorkflowAssociation);
                        if (!string.IsNullOrEmpty(loadedWorkflow?.Item1))
                        {
                            workFlowAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowDefinition(loadedWorkflow.Item1, WorkflowTypes.SP2010);
                        }

                        var workFlowTriggerAnalysisResult = WorkflowAnalyzer.Instance.ParseWorkflowTriggers(associatedWorkflow.WorkflowAssociation.AutoStartCreate, associatedWorkflow.WorkflowAssociation.AutoStartChange, associatedWorkflow.WorkflowAssociation.AllowManual);

                        results.Add(new WorkflowScanResult()
                        {
                            ListTitle = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.Title : "",
                            ListUrl = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.RootFolder.ServerRelativeUrl : "",
                            ListId = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.Id : Guid.Empty,
                            ContentTypeId = associatedWorkflow.AssociatedContentType != null ? associatedWorkflow.AssociatedContentType.StringId : "",
                            ContentTypeName = associatedWorkflow.AssociatedContentType != null ? associatedWorkflow.AssociatedContentType.Name : "",
                            Version = "2010",
                            Scope = associatedWorkflow.Scope,
                            RestrictToType = "N/A",
                            DefinitionName = associatedWorkflow.WorkflowAssociation.Name,
                            DefinitionDescription = "",
                            SubscriptionName = associatedWorkflow.WorkflowAssociation.Name,
                            HasSubscriptions = true,
                            Enabled = associatedWorkflow.WorkflowAssociation.Enabled,
                            DefinitionId = Guid.Empty,
                            IsOOBWorkflow = false,
                            SubscriptionId = associatedWorkflow.WorkflowAssociation.Id,
                            UsedActions = workFlowAnalysisResult?.WorkflowActions,
                            ActionCount = workFlowAnalysisResult != null ? workFlowAnalysisResult.ActionCount : 0,
                            UsedTriggers = workFlowTriggerAnalysisResult?.WorkflowTriggers,
                            LastSubscriptionEdit = associatedWorkflow.WorkflowAssociation.Modified,
                            LastDefinitionEdit = loadedWorkflow != null ? loadedWorkflow.Item2 : associatedWorkflow.WorkflowAssociation.Modified,
                            AssociationData = associatedWorkflow.WorkflowAssociation.AssociationData,
                            AllowManual = associatedWorkflow.WorkflowAssociation.AllowManual,
                            AutoStartChange = associatedWorkflow.WorkflowAssociation.AutoStartChange,
                            AutoStartCreate = associatedWorkflow.WorkflowAssociation.AutoStartCreate,
                            WorkflowTemplateName = associatedWorkflow.WorkflowAssociation.Name,
                            WFID = Guid.Empty.ToString(),
                            WebId = web.Id.ToString(),
                            WebUrl = web.Url,
                            SiteURL = cc.Site.Url,
                            SiteCollID = cc.Site.Id.ToString()
                        });
                    }
                }
                #endregion

            }
            catch(Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }

            return results;
        }



        private bool GetWorkflowPropertyBool(IDictionary<string, string> properties, string property)
        {
            if (string.IsNullOrEmpty(property) || properties == null)
            {
                return false;
            }

            if (properties.ContainsKey(property))
            {
                if (bool.TryParse(properties[property], out bool parsedValue))
                {
                    return parsedValue;
                }
            }

            return false;
        }

        private DateTime GetWorkflowPropertyDateTime(IDictionary<string, string> properties, string property)
        {
            if (string.IsNullOrEmpty(property) || properties == null)
            {
                return DateTime.MinValue;
            }

            if (properties.ContainsKey(property))
            {
                if (DateTime.TryParseExact(properties[property], "M/d/yyyy h:m:s tt", new CultureInfo("en-US"), DateTimeStyles.AssumeUniversal, out DateTime parsedValue))
                {
                    return parsedValue;
                }
            }

            return DateTime.MinValue;
        }

        private string GetWorkflowProperty(WorkflowSubscription subscription, string propertyName)
        {
            if (subscription.PropertyDefinitions.ContainsKey(propertyName))
            {
                return subscription.PropertyDefinitions[propertyName];
            }

            return "";
        }

        private Tuple<string, DateTime> LoadWorkflowDefinition(ClientContext cc, WorkflowAssociation workflowAssociation)
        {
            // Ensure the workflow library was loaded if not yet done
            LoadWorkflowLibrary(cc);
            try
            {
                  return GetFileInformation(cc.Web, $"{this.workflowList.RootFolder.ServerRelativeUrl}/{workflowAssociation.Name}/{workflowAssociation.Name}.xoml");
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
            }

            return null;
        }

        private Tuple<string, DateTime> LoadWorkflowDefinition(ClientContext cc, WorkflowTemplate workflowTemplate)
        {
            if (!IsOOBWorkflow(workflowTemplate.Id.ToString()))
            {
                // Ensure the workflow library was loaded if not yet done
                LoadWorkflowLibrary(cc);
                try
                {
                   return GetFileInformation(cc.Web, $"{this.workflowList.RootFolder.ServerRelativeUrl}/{workflowTemplate.Name}/{workflowTemplate.Name}.xoml");
                }
                catch (Exception ex)
                {
                    Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                    Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);
                }
            }

            return null;
        }

        private List LoadWorkflowLibrary(ClientContext cc)
        {
            if (this.workflowList != null)
            {
                return this.workflowList;
            }

            var baseExpressions = new List<Expression<Func<List, object>>> { l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, l => l.Hidden, l => l.RootFolder.ServerRelativeUrl };
            var query = cc.Web.Lists.IncludeWithDefaultProperties(baseExpressions.ToArray());
            var lists = cc.Web.Context.LoadQuery(query.Where(l => l.Title == "Workflows"));
            cc.ExecuteQuery();
            this.workflowList = lists.FirstOrDefault();

            return this.workflowList;
        }

        private static Tuple<string, DateTime> GetFileInformation(Web web, string serverRelativeUrl)
        {
            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);

            web.Context.Load(file);
            web.Context.ExecuteQuery();

            // TODO: fails when using sites.read.all role on xoml file download (access denied, requires ACP permission level)
            ClientResult<Stream> stream = file.OpenBinaryStream();
            web.Context.ExecuteQuery();

            string returnString = string.Empty;
            DateTime date = DateTime.MinValue;

            //UPDATED the LOC from date = file.ListItemAllFields.LastModifiedDateTime(); TO date = file.TimeLastModified;
            date = file.TimeLastModified;

            using (Stream memStream = new MemoryStream())
            {
                CopyStream(stream.Value, memStream);
                memStream.Position = 0;
                StreamReader reader = new StreamReader(memStream);
                returnString = reader.ReadToEnd();
            }

            return new Tuple<string, DateTime>(returnString, date);
        }

        private static void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

        private bool IsOOBWorkflow(string workflowTemplateId)
        {
            if (!string.IsNullOrEmpty(workflowTemplateId))
            {
                foreach (var oobId in WorkflowDiscovery.OOBWorkflowIDStarts)
                {
                    if (workflowTemplateId.StartsWith(oobId))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}

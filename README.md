# SP Workflow Scanner

This tool is written to scan existing SharePoint Server farms for SP 2010 and SP 2013 workflow associations and report on their composition (template, actions, etc) and a variety of metadata about the workflow. This tool should not be used for SharePoint Online (SPO). For SPO, check out the [Microsoft 365 Assessment tool](https://pnp.github.io/pnpassessment/index.html) from the PnP community.

## Install

We're working on publishing this package in PowerShell gallery. Once published, you can install the module using `Install-Module`. Until then, you will need to follow the [local build process](#build-process) below to clone and build the solution.

## Execution

Once the __SPWorkflowScanner__ module is imported, you can run a scan against a single site collection, multiple site collections (specified in a CSV), a single web application containing multiple site collections, or an entire SharePoint farm.

Our most recent tests have been performed against a SP2016 farm. Depending on what you want to scan, you may have to run the commands on a SharePoint server in  the farm itself. __This is necessary if you want to scan the Farm or a Web Application, as it requires access to SharePoint Server Object Model (SSOM) APIs.__ If only you want to scan a single site collection (or site collections in a CSV), these can be scanned from another computer that has network connectivity to the SharePoint WFE with SharePoint Client Side Object Model (CSOM) API calls.

This utility produces 4 folders, but only Logs and Summary have contents from this utility.

* Logs – contains the process logs for the utility and any errors.
* Summary – contains the WorkflowDiscovery.csv that has workflow information for all the site collections that were supplied in the sites.csv file.

![Folder_Structure](https://user-images.githubusercontent.com/63272213/137014648-a9ce8eb4-6e00-4bdd-aa39-2dde31a412a0.png)

The resulting WorkflowDiscovery.csv file should have the following columns for each workflow discovered:
| Column Name | Description |
| --- | --- |
| SiteURL | Site URL where workflow was found |
| SiteCollID | GUID of the site collection where workflow was found |
| WebURL | URL  |
| ListTitle | Title of list to which workflow is attached (if applicable) |
| ListUrl | URL of list to which workflow is attached (if applicable) |
| ContentTypeId | Content Type ID to which workflow is attached (if applicable) |
| ContentTypeName | Content Type Name to which workflow is attached (if applicable) |
| Scope | Scope of workflow (i.e. list, site, content type) |
| Version | 2010 or 2013 workflow |
| WFTemplateName | Template on which workflow is based (if applicable) Example: _Approval - SharePoint 2010_ |
| WorkFlowName | Name of workflow |
| IsOOBWorkflow | True/false value that indicates if the workflow is an out-of-the-box (OOB) workflow |
| Enabled | True/false value that indicates if workflow is enabled |
| WFID | GUID of workflow association |
| WebID | GUID of site/subsite where workflow was found |
| HasSubscriptions | True/false value that indicates if the workflow association has subscriptions |
| UsedActions | Semi-colon delimited list of actions in a SharePoint Designer (SPD) workflow |
| ToFlowMappingPercentage | Percentage of actions that map to Power Automate (formerly Flow). _-1_ if OOB or actions could not be scanned. |
| ConsiderUpgradingToFlow | True/false value if recommended to transform to Power Automate. This is solely based on if a workflow is enabled and has subscriptions. Additional review is required. |
| ActionCount | Count of actions in a SPD workflow |
| AllowManual | True/false if workflow allows manual trigger |
| AutoStartChange | True/false if workflow automatically starts on item change |
| AutoStartCreate | True/false if workflow automatically starts on item creation |
| LastDefinitionModifiedDate | Last modified date for workflow association __defintion__. Example: _3/28/2024 2:32:29 PM_ |
| LastSubscriptionModifiedDate | Last modified date for workflow association __subscription__. Example: _3/28/2024 2:32:29 PM_ |
| AssociationData | XML representation of SP workflow association |

### Examples

| Scenario  | Example |
| ------- | ---- |
| Using credentials of current logged in user, scan the __single__ Test123 __site collection__. | `Get-SPWorkflowScannerWorkflows -SiteCollectionUrl "https://sp2016.ad.xyz123domain.com/sites/Test123"` |
| Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain,scan the __single__ Test123 __site collection__. | `Get-SPWorkflowScannerWorkflows -SiteCollectionUrl "https://sp2016.ad.xyz123domain.com/sites/Test123" -Credential $Cred -Domain "ad.xyz123domain.com"`|
| Using credentials of current logged in user, scan __multiple site collections__ in the sites.csv file. _Each line in sites.csv is a URL for a site collection. No header line is present._ | `Get-SPWorkflowScannerWorkflows -SiteCollectionURLFilePath .\sites.csv` |
| Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain, scan __multiple site collections__ in the sites.csv file. _Each line in sites.csv is a URL for a site collection. No header line is present._ | `Get-SPWorkflowScannerWorkflows -SiteCollectionURLFilePath .\sites.csv -Credential $Cred -Domain "ad.xyz123domain.com"` |
| Using a PSCredential object ($Cred), enumerate all sites in the provided __web application__ and scan for workflow in the site collections. | `Get-SPWorkflowScannerWorkflows -WebApplicationUrl "https://sp2016.ad.xyz123domain.com"` |
| Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain, enumerate all sites in the provided __web application__ and scan for workflow in the site collections. | `Get-SPWorkflowScannerWorkflows -WebApplicationUrl "https://sp2016.ad.xyz123domain.com" -Credential $Cred -Domain "ad.xyz123domain.com"` |
| Using credentials of current logged in user, enumerate all sites in the __farm__. | `Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput"` |
| Using a PSCredential object ($Cred) for authentication on the ad.xyz123domain.com domain, enumerate all sites in the __farm__. | `Get-SPWorkflowScannerWorkflows -AssessmentOutputFolder "C:\AssessmentOutput" -Credential $Cred -Domain "ad.xyz123domain.com"` |

#### Additional notes

* The account/credential provided must have the equivalent of __site collection administrator__ access to each site collection to be scanned.
*If scanning the web application, the account requires full control on the web application as well.
* If scanning the entire farm, the account must be a farm administrator as well.
* See note above regarding from what computer/server you need to run on dependent on your scenario.

## Build process

Clone Project locally. This project has been verified to work in Visual Studio 2022. Once project is open, restore NuGet packages and you should be able to build the solution.

There is a MSTest unit test project in the solution for testing the WorkflowScanner library and the PowerShell module. __To run tests targeting a web application or farm, you must run the tests on a SharePoint Server in the farm.__ The tests rely on configuration values in the app.config of the Test project for targeting specific web applications and site collections, for specifying credentials to use for authentication in tests passing an explicit login, and the expected findings for each test (dependent on your farm and environment). These tests cannot be run without a SharePoint farm to scan.

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

## Origins

This base code was developed by the Modern Work Team from the [Industry Solutions Delivery group](https://www.microsoft.com/en-us/msservices/modern-work) at Microsoft.

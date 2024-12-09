# Introduction

Oversharing script is PowerShell based solution that allows an M365
tenant admin/ SPO Administrator to report on documents that are
overshared. This solution scans SharePoint Online as well as OneDrive
for Business (ODFB).

Below are the sharing scenarios that this script reports on:

| Target Area   | Scenario |
| ------------- | ------------- |
| Site  | Shared with the default SPO "Everyone", "Everyone except external" or security group which has everyone within the organization  |
| Site  | Viewable by external users  |
| Site  | Viewable by anonymous users  |
| Library  | Shared with the default SPO "Everyone", "Everyone except external" or security group which has everyone within the organization  |
| Library  | Viewable by external users  |
| Library  | Viewable by anonymous users  |
| Item  | Shared with the default SPO "Everyone", "Everyone except external" or security group which has everyone within the organization  |
| Item  | Accessible via "Anyone" sharing link  |
| Item  | Viewable by external users  |

This document describes the steps required to deploy and verify the
installation and execution of Oversharing Script.

## Oversharing script architecture

![](media/SolutionArchitecture.png)

Oversharing script is a simple implementation to discover overshared
information within an M365 tenant. The script uses SharePoint online
native search Apis to query data from SPO/ODFB and outputs in CSVs.
Since the implementation utilizes SPO Search & query-based engine:

a.  Script only can process information which is made searchable within
    a tenant. It wont surface information from sites for which Search
    has been disabled i.e.

![A screenshot of a computer Description automatically
generated](media/SPOSearchConfiguration.png){width="5.647833552055993in"
height="1.6622287839020122in"}

Fig. Site Collection Search Setting

![](media/SPOSearchLibConfiguration.png){width="3.7916666666666665in"
height="3.2083333333333335in"}

Fig. SharePoint Library Search Setting

b.  Script can be customized to target different scenarios for ex: only
    discover information from specific sites, sites with specific
    metadata attached to it, specific file types & others

## Intended Audience

This document is intended for the team responsible for managing M365 artefacts, such as SharePoint Online (SPO)
sites & One Drive for Business (ODFB). It assumes a basic working knowledge of M365, SharePoint Online, ODFB, PowerShell and PnP.PowerShell.

# Deployment Guidance

The following sections document the steps required to deploy
"Oversharing Solution".

## Prerequisites

This solution has various prerequisites that should be confirmed as
ready before installation if possible. Pre-requisites are below.

### Machine (If running from a local server)

If the script is running from a server, then, a machine with the
hardware and software configuration below is required:

1.  **Hardware Configuration**<br />
    a. Processor: 1 gigahertz (GHz) or fasterRAM<br />
    b. 4 gigabytes (GB)Storage<br />
    c. 64 GB
2.  **Software Requirements**
    a. OS: Windows 10 1607+, Windows 11, Windows Server 2016, 2019, 2022

3.  Good internet connectivity

### PowerShell 7.2+

This solution uses the latest **PnP.PowerShell** PowerShell module which
requires PowerShell 7.2+. More Info
<https://pnp.github.io/powershell/articles/installation.html>

### PowerShell Modules

The below PowerShell modules are required:


> Install the PnP.PowerShell Module (min v2.3)Install-Module
> PnP.PowerShell -MinimumVersion 2.3.0

### Required User Account

Please create a dummy user account (**[No M365 license assignment
required]{.underline}**) in Azure Entra which we will be using to run
the oversharing script to scan and report on items which are accessible
to everyone with organisation.

**[NOTE:]{.underline}** Make sure this user account is not given access
to any M365 resources.

### Authentication

The solutions use latest Microsoft Entra App-only authentication
architecture to authenticate against the SPO/ODFB when retrieving data
which is accessible to externals/anonymous users. ([More
Information](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread)).

**[Note: If reporting on SPO resources accessible to externals/anonymous
is not required then we don't need to do below mentioned
steps.]{.underline}**

#### Microsoft Entra App Registration & Permission configuration

-   

> Navigate to <https://entra.microsoft.com/> and click on "App
> Registration".![A screenshot of a computer Description automatically
> generated](media/image4.png){width="1.3721412948381453in"
> height="3.3116119860017497in"}

-   

> Click "New Registration" and fill in the details and click "Register".
>
> ![](media/image5.png){width="4.7985400262467195in"
> height="1.9183902012248468in"}
>
> ![](media/image6.png){width="4.729166666666667in" height="4.53125in"}

-   

> *[Copy the "Application ID" as we would need this for our Oversharing
> Script execution.]{.underline}*![](media/image7.png){width="4.4375in"
> height="2.0520833333333335in"}

-   

> Navigate to "API Permissions", "Add a permission", and then
> "SharePoint"
>
> ![](media/image8.png){width="5.813531277340332in"
> height="2.8229166666666665in"}

-   Click on "Application permissions" 🡪 "Sites.Read.All" and click "Add permission"
>
> ![](media/image9.png){width="4.375in" height="4.46875in"}

-   Click "Grant admin Consent for ......" and complete the consent process.Your API permissions should now display as below:

![](media/image10.png){width="4.575694444444444in"
height="0.5597222222222222in"}

#### Microsoft Entra App certificate

When wanting to use certificate-based authentication for M365 Entra App
to connect to SPO/ODFB, follow steps defined in below link to generate a
self-signed certificate.

[Granting access via Azure AD
App-Only](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread)

Once certificate is generated, upload the "**.cer**" file to the App
registered in step#2.2.1.1

![](media/image11.png){width="4.677083333333333in"
height="3.6458333333333335in"}

## Oversharing script execution 

Oversharing script has 2 operation modes which are mentioned below:

-  Scan all SPO resources accessible to everyone within the Organisation
-  Scan all SPO resources accessible to anonymous and external users


### Scan all SPO resources accessible to everyone within the Organisation

For this scenario, oversharing script will require below mandatory
information

  ----------------------------------------------------------------------------------------------------------------------------
  Parameter Name                      Details
  ----------------------------------- ----------------------------------------------------------------------------------------
  SPTenantName                        SharePoint Tenant Name. If you SPO Admin Url is
                                      [https://[contoso]{.mark}-admin.sharepoint.com](https://contoso-admin.sharepoint.com),
                                      then tenant name is "Contoso"

  PnPEnterpriseAppId                  Provide the PnP Enterprise App ID as per the new authentication model added in
                                      PnP.PowerShell module when using interactive login ([more
                                      info](https://pnp.github.io/powershell/articles/registerapplication))

  ScanAccessibleToEveryone            Switch to specify scanning to be done for all data accessible to everyone within
                                      organization

  outputreportpath                    Local path on the machine where the reports will be downloaded

  Logpath                             Local path on the machine to output the script execution log
  ----------------------------------------------------------------------------------------------------------------------------

**[Example:]{.underline}**

.\\GenerateOversharedDataReportV3.ps1 -SPTenantName \"contoso\"
-TenantID \"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\"
 -ScanAccessibleToEveryone -PnPEnterpriseAppId
\"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\" -outputreportpath
\"c:\\sharingreport\\\" -logpath \"c:\\sharingreport\\\"

### Scan all SPO resources accessible to anonymous and external users

For this scenario, oversharing script will use M365 AppID and
certificate thumbprint to connect to SPO/ODFB to discover spo resources
accessible to external/anonymous users. To use certificate-based
authentication to connect to SPO/ODFB, make sure the ".pfx" file
generated during step#2.1.5.2 is installed on the local machine from
where the oversharing script will be executed. To install the ".pfx"
file, follow below steps

-   

Double click ".pfx" file
![](media/image12.png){width="3.065825678040245in"
height="2.4799715660542434in"}

![](media/image13.png){width="2.752083333333333in"
height="2.7118055555555554in"}

-   

Enter the password that was set at the time of certificate
generation.![](media/image14.png){width="2.7756944444444445in"
height="2.6479166666666667in"}

![](media/image15.png){width="2.736111111111111in"
height="2.6083333333333334in"}

![](media/image16.png){width="2.7756944444444445in"
height="2.7041666666666666in"}

For above scenario, oversharing script will require below mandatory
information

  ----------------------------------------------------------------------------------------------------------------------------
  Parameter Name                      Details
  ----------------------------------- ----------------------------------------------------------------------------------------
  SPTenantName                        SharePoint Tenant Name. If you SPO Admin Url is
                                      [https://[contoso]{.mark}-admin.sharepoint.com](https://contoso-admin.sharepoint.com),
                                      then tenant name is "Contoso"

  ClientID                            M365 App Entra ID created in step#2.1.5.1

  Thumbprint                          M365 App Certificate thumbprint created in step#2.1.5.2

  TenantID                            M365 Tenant ID

  outputreportpath                    Local path on the machine where the reports will be downloaded

  Logpath                             Local path on the machine to output the script execution log
  ----------------------------------------------------------------------------------------------------------------------------

**[Example:]{.underline}**

.\\GenerateOversharedDataReportV3.ps1 -SPTenantName \"contoso\"
-TenantID \"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\"
 -ScanAccessibleToExternalsAndAnonymous -ClientID
\"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\" -Thumbprint
\"911D70A3C206A14EDC5342D9EDAJSGHEEHJ\" -outputreportpath
\"c:\\sharingreport\\\" -logpath \"c:\\sharingreport\\\"

# FAQs

1)  Can we run schedule this script to run using Azure Automation?Yes, this
script can be scheduled to run using Azure automation with minor tweaks.

2)  Does this script output all the information into single CSV file?No, the
script is designed to generate multiple CSVs files and each CSV by
default will have 10,000 rows of data (this value can be changed if
number of rows needed per csv \>10,000).

3)  How to fast track the data discovery process using the script?The best
possible approach is to run the script in parallel targeting different
areas within SPO/ODFB to speed up the overshared data discovery process.
For example, one script can scan only the most sensitive sites that you
may have, another can scan just ODFB and so on.

4)  How to target the script to just scan specific areas within SharePoint
or ODFB?Since the oversharing script uses SharePoint KQL search queries,
it gives us flexibility to tailor the query to limit scanning to
designated areas within SPO. One would need KQL skills to build the
query.

5)  Can I extract additional metadata from the oversharing script?Yes, if
the metadata that needs to be extracted is crawled by SPO Search APIs
and is available, we should be able to extract that into the output
CSVs.

6)  Do we know how much time the script will take to report on
oversharing?The script is highly efficient but it's difficult to
estimate how much time it will take. Since each tenant is different in
terms of size, it becomes increasingly difficult to estimate on time it
will take. There are other factors that affect the performance of the
script i.e.<br />
    a.  Internet connection speed<br />
    b.  Network configuration<br />
    c.  Number of metadata

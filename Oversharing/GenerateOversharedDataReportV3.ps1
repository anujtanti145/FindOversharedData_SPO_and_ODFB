# DiSCLAIMER:
# This script is made available to you without any express, implied or
# statutory warranty, not even the implied warranty of
# merchantability or fitness for a particular purpose, or the
# warranty of title or non-infringement. The entire risk of the
# use or the results from the use of this script remains with you

<#
.SYNOPSIS
This is Powershell script generates report on overshared files and resources accessible to externals across SPO Tenant and saves the report locally on machine where script is executed from.

.DESCRIPTION
This is Powershell script generates report on overshared files and resources accessible to externals across SPO Tenant and saves the report locally on machine where script is executed from.

.PARAMETER SPTenantName
Provide the SPTenantName Name i.e. for contoso.sharepoint.com provide contoso

.PARAMETER ScanAccessibleToEveryone
Switch to specify if scan to be done for all data accessible to everyone within organization

.PARAMETER PnPEnterpriseAppId 
Provide the PnP Enterprise App ID as per the new authentication model added in PnP.PowerShell module when using interactive login (more info:https://pnp.github.io/powershell/articles/registerapplication)
This is required mandatory parameter when ScanAccessibleToEveryone switch is provided

.PARAMETER ScanAccessibleToExternalsAndAnonymous
Switch to specify if scan to be done for all data accessible to external users and anonymous users

.PARAMETER ClientID
Provide the Azure Entra App ID that is authorized to read all data from SPO & ODFB.
This is mandatory parameter when ScanAccessibleToExternalsAndAnonymous switch is provided

.PARAMETER Thumbprint 
Provide the Azure Entra App certificate thumbprint registered against the param ClientID for authentication
Mandatory parameter when ScanAccessibleToExternalsAndAnonymous switch is provided

.PARAMETER TenantID
Provide M365 TenantID

.PARAMETER outputreportpath
Provide output report file path

.PARAMETER logpath
Provide the path to save the logs

.PARAMETER dataRowCountPerCSV
Provide the number of rows to be saved in each CSV file. Default value is 10000


.EXAMPLE
./GenerateOversharedDataReportV3.ps1 -SPTenantName "contoso" -TenantID "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"  -ScanAccessibleToExternalsAndAnonymous -ClientID "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Thumbprint "911D70A3C206A14EDC5342D9EDAJSGHEEHJ" -outputreportpath "c:\sharingreport\" -logpath "c:\sharingreport\"

./GenerateOversharedDataReportV3.ps1 -SPTenantName "contoso" -TenantID "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"  -ScanAccessibleToEveryone -PnPEnterpriseAppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -outputreportpath "c:\sharingreport\" -logpath "c:\sharingreport\"

.NOTES
   Author: Anuj Tanti
   Created: 15/11/2023
   Last Modified: 18/06/2024
   Version: 2.0  
#>

Param
(
    [Parameter(Mandatory = $true)]
    [string] $SPTenantName,

    [Parameter(Mandatory = $true, ParameterSetName = "ScanEverything")]
    [switch] $ScanAccessibleToEveryone,

    [Parameter(Mandatory = $true, ParameterSetName = "ScanEverything")]
    [string] $PnPEnterpriseAppId,

    [Parameter(Mandatory = $true, ParameterSetName = "ScanExternalSharing")]
    [switch] $ScanAccessibleToExternalsAndAnonymous,

    [Parameter(Mandatory = $true, ParameterSetName = "ScanExternalSharing")]
    [ValidatePattern("^\{?[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}\}?$")]
    [string] $ClientID,
    
    [Parameter(Mandatory = $true, ParameterSetName = "ScanExternalSharing")]
    [string] $Thumbprint,
    
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^\{?[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}\}?$")]
    [string] $TenantID,
    
    [Parameter(Mandatory = $true)]
    [string] $logpath,

    [Parameter(Mandatory = $true)]
    [string] $outputreportpath,

    [int] $dataRowCountPerCSV = 10000
  
)


function CheckPnPPowerShellModuleInstallation() {
    $PnPModule = Get-InstalledModule -Name PnP.PowerShell
    
    if (-not $PnPModule) {
        Write-Host -f Cyan "PnPPowerShell Module not found..!!!"
        Write-Host ""

        Write-Host "Installing PnP.PowerShell Module"
        Install-Module PnP.PowerShell -Force
    }
    else {
        Write-Host ""
        Write-Host -f Green "PnPPowerShell Module found..Continuing next steps..!!!"
        Write-Host ""
    }

}

function CheckPowerShellVersion() {
    $PSversion = $PSVersionTable.PSVersion | Select-Object Major
    $isPS7 = $true
    if ($PSversion.Major -ne "7") {
        Write-Host -f Cyan "This script requires PowerShell v7!!!"
        Write-Host ""
        $isPS7 = $false
    }
    return $isPS7
}


function GetOverSharedDataForSiteCollection([int]$sIndex, [int]$rLimit, $srcQuery) {
    Execute-WithRetry {
        $searchResults = Submit-PnPSearchQuery -Query $srcQuery  -StartRow $sIndex -MaxResults $rLimit -TrimDuplicates $false -SortList @{"LastModifiedTime" = "descending" } -SelectProperties SitePath, UniqueId, contentclass, SiteTemplate, Path, FileType, IsDocument, IsContainer, LastModifiedTime, CreatedBy, ModifiedBy,ChannelGroupId, GroupId, RelatedGroupId,ViewableByAnonymousUsers, ViewableByExternalUsers, InformationProtectionLabelId, AuthorOWSUser, EditorOWSUser
        return $searchResults 
    }
}


function GetSearchQueryBasedOnScanMode() {
    $srcQuery = '(contentclass:STS_Site OR contentclass:STS_Web OR (contentclass:STS_ListItem* AND (IsContainer:1 OR IsDocument:1) AND FileType<>aspx) OR contentclass:STS_List_*) AND -SiteTemplate:APPCATALOG'
    if ($ScanAccessibleToExternalsAndAnonymous) {
        $srcQuery = "ViewableByExternalUsers:1 OR ViewableByAnonymousUsers:1"
    }
    return $srcQuery
}


Function ConnectToTenant($SPRootSitUrl) {
    if ($ScanAccessibleToEveryone) {
        # Display the message box and wait for the user to click OK
        [void][System.Windows.Forms.MessageBox]::Show("Please login with the temporary user account that has been created for oversharing scan", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        Connect-PnPOnline -Url $SPRootSitUrl -Interactive -ClientId $PnPEnterpriseAppId -WarningAction Ignore
    }
    else {
        Connect-PnPOnline -Url $SPRootSitUrl -ClientId $ClientID -Thumbprint $Thumbprint -Tenant $TenantID -WarningAction Ignore
    }
} 

function IdentifyContentTypeShared($contentclass) {
    $contentType = ""
    switch -wildcard ($contentclass) {
        "*STS_List_*" { $contentType = "SharePoint List or DocLibrary" }
           
        "*STS_ListItem*" { $contentType = "SharePoint List Item" }
           
        "STS_Site" { $contentType = "SharePoint Site Collection" }
           
        "STS_Web" { $contentType = "SharePoint Subsite" }
           
    }
    
    Return $contentType
}

function AddDataToCollection($resultCollection) {
    $results = @()
    
    foreach ($ResultRow in $resultCollection.ResultRows) {      
        $Result = [PSCustomObject]@{
            SiteUrl                  = $ResultRow["SitePath"]
            SPOResourceType          = IdentifyContentTypeShared $ResultRow["contentclass"]
            SiteTemplate             = $ResultRow["SiteTemplate"]
            IsDocument               = $ResultRow["IsDocument"]
            IsFolder                 = $ResultRow["IsContainer"]
            ItemPath                 = $ResultRow["Path"]
            DocumentUniqueID         = $ResultRow["UniqueId"]
            FileType                 = $ResultRow["FileType"]
            LastModifiedTime         = $ResultRow["LastModifiedTime"]
            CreatedBy                = ($null -ne $ResultRow["AuthorOWSUser"])? ($ResultRow["AuthorOWSUser"].ToLower().Contains("system account") -or $ResultRow["AuthorOWSUser"].ToLower().Contains("sharepoint app")) ? $ResultRow["AuthorOWSUser"].Split("|")[1]:$ResultRow["AuthorOWSUser"].Split("|")[0] :$ResultRow["CreatedBy"] 
            ModifiedBy               = ($null -ne $ResultRow["EditorOWSUser"])? ($ResultRow["EditorOWSUser"].ToLower().Contains("system account") -or $ResultRow["EditorOWSUser"].ToLower().Contains("sharepoint app")) ? $ResultRow["EditorOWSUser"].Split("|")[1]:$ResultRow["EditorOWSUser"].Split("|")[0] :$ResultRow["ModifiedBy"] 
            GroupId                  = $ResultRow["GroupId"]
            RelatedGroupId           = $ResultRow["RelatedGroupId"]
            ChannelGroupId           = $ResultRow["ChannelGroupId"]
            ViewableByAnonymousUsers = $ResultRow["ViewableByAnonymousUsers"]
            ViewableByExternalUsers  = $ResultRow["ViewableByExternalUsers"]
            SensitivityLabelID       = $ResultRow["InformationProtectionLabelId"]
        }
        $results += $Result
    }

    return $results
}


function main() {
    Add-Type -AssemblyName System.Windows.Forms
    $logfileName = "log" + (Get-Date -f "ddMMyyyy_hhmm") + ".txt"
        
    Try {
        Start-Transcript -Path ($logpath + "/" + $logfileName) -ErrorAction Stop
        Write-Host ""
    }
    catch {
        Start-Transcript -Path ($logpath + "/" + $logfileName)
    }
    
    if (CheckPowerShellVersion) {

        $starttime = Get-Date
        Write-Host "Start Time $($starttime)"
        Write-Host ""
        $SPRootSitUrl = "https://$($SPTenantName).sharePoint.com/"

        ConnectToTenant -SPRootSitUrl $SPRootSitUrl

        $rowLimit = (500 -ge $dataRowCountPerCSV) ? $dataRowCountPerCSV : 500 #($rowLimit = 500 #default limit and cannot be greater than 500)
        $startIndex = 0
        $sharedDataCollection = @()       
        $csvStartRowCount = 1
        $csvEndRowCount = $dataRowCountPerCSV
        $srcQuery = GetSearchQueryBasedOnScanMode
        
        do {

            $resultCollection = GetOverSharedDataForSiteCollection -sIndex $startIndex -rLimit $rowLimit -srcQuery $srcQuery
        
            $sharedDataCollection += AddDataToCollection -resultCollection $resultCollection

            $startIndex += $resultCollection.RowCount
            Write-Host ""
            if ($ScanAccessibleToEveryone) {
                Write-Host -f Green "Processed $($startIndex) records out of Total $($resultCollection.TotalRows) overshared records"
            }
            else {
                Write-Host -f Green "Processed $($startIndex) records out of Total $($resultCollection.TotalRows) externally accessible records"
            }
            
            #this logic will generate multiple CSVs based the number of rows desired per CSV for large tenants
            if (($sharedDataCollection.Count -ge $dataRowCountPerCSV) -or ($startIndex -eq $resultCollection.TotalRows) -or ($resultCollection.RowCount -eq 0)) {
                if ($startIndex -eq $resultCollection.TotalRows) {
                    $csvEndRowCount = $resultCollection.TotalRows
                }
                $sharingReportFileName = "SharingReport_$($csvStartRowCount)_$($csvEndRowCount)_" + (Get-Date -f "ddMMyyyy_hhmm") + ".csv"
                
                $sharedDataCollection | Export-Csv -Path ($outputreportpath + "\$($sharingReportFileName)")
                              
                $csvStartRowCount += $sharedDataCollection.Count
                $csvEndRowCount += $dataRowCountPerCSV
                $sharedDataCollection = @()
            }

        }while ($resultCollection.TotalRows -gt $startIndex) 

        Write-Host ""
        Write-Host "Total Execution Time in Seconds $((New-TimeSpan -Start $starttime -End (Get-Date)).TotalSeconds)"
        
    }
    Write-Host ""
    Stop-Transcript
    Write-Host ""
}

function Execute-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$ScriptBlock,
        [Parameter(Mandatory = $false)]
        [int]$MaxRetryCount = 5
    )

    $retryCount = 0
    while ($true) {
        try {
            & $ScriptBlock
            return
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503 -or $_ -like "*Timeout*" -or $_ -like "*Search has encountered a problem*") {
                if ($retryCount -ge $MaxRetryCount) {
                    throw
                }
                else {
                    $retryAfter = $null
                    if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
                        $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                    }
                    if ($null -ne $retryAfter) {
                        Start-Sleep -Seconds $retryAfter
                    }
                    else {
                        Write-Host "Error occurred: $($_.Exception.Message). Retrying... ($retryCount/$MaxRetryCount)" -ForegroundColor Yellow
                        Start-Sleep -Seconds 5 # Optional: Wait for 5 seconds before retrying
                    }
                    $retryCount++
                }
            }
            else {
                throw
            }
        }
    }
}

# Check and Install PnPPowerShell Module if required
CheckPnPPowerShellModuleInstallation

#start main function
main

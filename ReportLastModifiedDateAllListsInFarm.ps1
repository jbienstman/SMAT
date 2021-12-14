<#
.NOTES
    #######################################################################################################################################
    # Author: Jim B.
    # Latest version can be found on Github: https://github.com/jbienstman/SMAT
    #######################################################################################################################################
    # Revision(s)
    # 1.0.0 2021-10-14  Initial Version
    # 1.1.0 2021-11-08  Updated script cosmetics
    # 1.2.0 2021-11-08  Now checking whether the $reportPath exists (added Function: AskYesNoQuestion)
    # 1.2.1 2021-11-18  Updated Header Synopsis, Description, Example ...
    #######################################################################################################################################

.SYNOPSIS
    On a SharePoint Server 2010+ Farm (On-Premise), this script will iterate through all lists in all webs in all sites in all web applications
    and output a CSV report with the Last Modified Date and Modified By (if enabled and available) for each list.
    This is meant to be more accurate than the web level property "LastItemModifiedDate" provided by default.

.DESCRIPTION
    Prerequisites:
    - The account used to run the script must have the following permissions:
        - SharePoint Farm Administrator
        - SPShellAdmin for all content databases: https://docs.microsoft.com/en-us/powershell/module/sharepoint-server/add-spshelladmin
        - Local Administrator permissions on the SharePoint server where you will run this script
        - Minimum "full read" in all web application user policies: https://docs.microsoft.com/en-us/sharepoint/administration/manage-permission-policies-for-a-web-application#add-users-to-or-remove-users-from-a-permission-policy-level
    - SharePoint PowerShell Snap-in: "Microsoft.SharePoint.PowerShell" (*)

    Running the Script:
    - You can run the script as a whole by updating the parameters directly or you can call the script from a PowerShell window (*)

.EXAMPLE
    ReportLastModifiedDateAllListsInFarm -reportPath "C:\TEMP\" -reportLastModifiedBy:$true -outputEnabled:true

#>
###########################################################################################################################################
Param (
    [Parameter(Mandatory=$false, HelpMessage = "Folder location where report should be saved")][string]$reportPath = "C:\TEMP\" ,
    [parameter(mandatory=$false, HelpMessage = "Run additional query to get identity of last modified by")][bool]$reportLastModifiedBy = $true ,
    [parameter(mandatory=$false, HelpMessage = "Output script progress to screen")][bool]$outputEnabled = $true ,
    [parameter(mandatory=$false, HelpMessage = "Enable PowerShell Transcripting")][bool]$EnableLogging = $false #NOTE: This logs nothing if $outputEnabled = $false!
)
###########################################################################################################################################
#region - TRY
try
{
#region - static variable(s)
$ScriptRequiresAdminPrivileges = $True
$ScriptRequiresSharePointAddin = $True
$ExcludedLists = @(
"Access Requests",
"App Packages",
"appdata",
"appfiles",
"Apps in Testing",
"Cache Profiles",
"Composed Looks",
"Content and Structure Reports",
"Content type publishing error log",
"Converted Forms",
"Device Channels",
"Form Templates",
"fpdatasources",
"Get started with Apps for Office and SharePoint",
"List Template Gallery",
"Long Running Operation Status",
"Maintenance Log Library",
#"Images",
#"site collection images",
"Master Docs",
#"Master Page Gallery",
"MicroFeed",
"NintexFormXml",
"Quick Deploy Items",
"Relationships List",
#"Reusable Content",
"Reporting Metadata",
"Reporting Templates",
"Search Config List",
#"Site Assets",
#"Pages",
#"Preservation Hold Library",
#"Site Pages",
"Solution Gallery",
#"Style Library",
"Suggested Content Browser Locations",
"Tabs in Search Pages",
"Tabs in Search Results",
"Theme Gallery",
"TaxonomyHiddenList",
"User Information List",
"Web Part Gallery",
#"Workflow History",
#"Workflow Tasks",
"wfpub",
"wfsvc"
)
#endregion - static variable(s)
###########################################################################################################################################
#region - Minimal Header - v1.0.0
#region - StartTime & Preferences
$startTime = (Get-Date)
if ($outputEnabled) {
    Clear-Host
    Write-Host ("Script Started: ") -NoNewline -ForegroundColor DarkCyan
    Write-Host $startTime -ForegroundColor DarkYellow
}
$ErrorActionPreference = "Stop";
#endregion - StartTime & Preferences
#region - ScriptPath
if ($outputEnabled) {Write-Host "Getting ScriptPath: " -NoNewline}
$ScriptPath = $PSScriptRoot
if ($MyInvocation.InvocationName.Length -ne "0")
    {
    #$ScriptPath = Split-Path $MyInvocation.InvocationName
    $ScriptFileName = $MyInvocation.MyCommand.Name
    $ScriptFileNameNoExtension = ($ScriptFileName.Split("."))[0]
    if ($outputEnabled) {Write-Host "OK" -ForegroundColor Green}
    }
else
    {
    if ($outputEnabled) {Write-Warning "Cannot get Script Path, stopping script..."}
    exit
    }
#endregion - ScriptPath
#region - Run As Admin
if ($ScriptRequiresAdminPrivileges)
    {
    Function IsAdmin
        {
        $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
        return $IsAdmin
        }
    if ($outputEnabled) {Write-Host "Run As Admin: " -NoNewline -ForegroundColor Gray}
    if((IsAdmin) -eq $false)
        {
        if ($outputEnabled) {
        Write-Host "NO" -ForegroundColor Red
        Write-Warning "This Script requires `"Administrator`" privileges, stopping..."
        }
        return
        }
    else
        {
        if ($outputEnabled) {Write-Host "OK" -ForegroundColor Green}
        }
    }
#endregion - Run As Admin
#region - SharePoint Snapin
if ($ScriptRequiresSharePointAddin)
    {
    if ($outputEnabled) {Write-Host "Loading SharePoint PowerShell Snapin: " -NoNewline -ForegroundColor Gray}
    if ($null -eq (Get-PSSnapin "Microsoft.SharePoint.PowerShell" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue))
        {
        Add-PSSnapin "Microsoft.SharePoint.PowerShell" -WarningAction SilentlyContinue -ErrorAction Stop
        if ($outputEnabled) {Write-Host "OK" -ForegroundColor Green}
        }
    else
        {
        if ($outputEnabled) {Write-Host "OK (Already Loaded)" -ForegroundColor Green}
        }
    }
#endregion - SharePoint Snapin
#region - EnableLogging
if ($EnableLogging) {Start-Transcript -Path ($ScriptPath + "\" + $ScriptFileNameNoExtension + "_" + (Get-Date -Format yyyyMMddHHmmss) + ".log")}
#endregion - EnableLogging
#endregion - Minimal Header - v1.0.0
###########################################################################################################################################
#region - Function(s)
Function AskYesNoQuestion {
    <#
    .EXAMPLE
        AskYesNoQuestion ("Your Question Text Here?")
    #>
    Param (
        [Parameter(Mandatory=$true)][string]$Question ,
        [Parameter(Mandatory=$false)][string]$ForegroundColor = "White",
        [Parameter(Mandatory=$false)][string]$Choice1 = "y" ,
        [Parameter(Mandatory=$false)][string]$Choice2 = "n"
    )
    $QuestionSuffix = "[$Choice1/$Choice2]"
    Do {Write-Host ($Question) -ForegroundColor $ForegroundColor -NoNewline;[string]$CheckAnswer = Read-Host $QuestionSuffix}
    Until ($CheckAnswer -eq $Choice1 -or $CheckAnswer -eq $Choice2)
    Switch ($CheckAnswer)
        {
            $Choice1 {Return $True}
            $Choice2 {Return $False}
        }
}
#endregion - Function(s)
###########################################################################################################################################
#region - Main
$spFarm = Get-SPFarm
$spWebApplications = Get-SPWebApplication
if ($outputEnabled)
    {
    $webAppWarningCount = 10
    if ($spWebApplications.Count -gt $webAppWarningCount)
        {
        $continue = AskYesNoQuestion ("WARNING: There are more than $webAppWarningCount Web Applications in your Farm - You still want to continue? ") -ForegroundColor Red
        if ($continue -eq $false)
            {
            exit
            }
        }
    }
if (Test-Path $reportPath.Trim("\"))
    {
    #$reportPath exists, script can continue
    }
else
    {
    #$reportPath does NOT exists, script halted
    Write-Warning ("The path: " + $reportPath + " does not exist!")
    exit
    }
$reportFullPathName = ($reportPath.Trim("\") + "\" + "LastModifiedDatesListsInFarm_" + $spFarm.ID.Guid + "_" + (Get-Date -Format yyyyMMddHHmmss) + ".csv")
$spObjects = @()
$i = 0
foreach ($spWebApplication in $spWebApplications)
    {
    $i++
    if ($outputEnabled) {Write-Host ("[" + $i + "/" + $spWebApplications.Count +  "] (web application) - `"" + $spWebApplication.Url + "`"") -foreGroundcolor Cyan}
    $spSites = $spWebApplication.Sites
    if ($spWebApplication.Sites.Count -gt 0)
        {
        $j = 0
        foreach ($spSite in $spSites)
            {
            $j++
            #GET CHANGES - SEPARATE REPORT
            if ($outputEnabled) {Write-Host (" - [" + $j + "/" + $spSites.Count +  "] (spSite) - `"" + $spSite.Url + "`"") -foreGroundcolor Yellow}
            $k = 0
            foreach ($spWeb in $spSite.AllWebs)
                {
                $k++
                if ($outputEnabled) {Write-Host (" - - + [" + $k + "/" + $spSite.AllWebs.Count +  "] (spWeb) - `"" + $spWeb.ServerRelativeUrl + "`"") -foreGroundcolor Gray}
                $spLists = $spWeb.Lists | Where-Object {$_.Hidden -eq $False -and $ExcludedLists -notcontains $_.Title}
                if ($null -eq $spLists)
                    {
                    if ($outputEnabled) {Write-Host (" - - - + NO LISTS FOUND! - SKIPPING WEB ") -foreGroundcolor Yellow}
                    }
                else
                    {
                    $l = 0
                    foreach ($spList in $spLists)
                        {
                        $l++
                        #region - check list item count
                        $ListItemCount = $spList.ItemCount
                        if ($ListItemCount -le 0)
                            {
                            $ListItemCount = $spList.Items.Count #double checking if the item count is suspiciously low
                            }
                        #endregion - check list item count
                        if ($outputEnabled) {Write-Host (" - - - + [" + $l + "/" + $spLists.Count +  "] (" + $spList.BaseTemplate + ") - " + $spList.Title + " (#" + $ListItemCount + ")") -foreGroundcolor DarkGray}
                        #region - query to get "last modified by" info
                        if ($reportLastModifiedBy)
                            {
                            if ($ListItemCount -ne 0)
                                {
                                $Query=New-Object Microsoft.SharePoint.SPQuery
                                $Query.RowLimit = 1
                                #$Query.Query = "<OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy>"
                                #$Query.ViewXml = "<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy></Query><RowLimit Paged=TRUE'>1</RowLimit></View>"
                                $Query.ViewXml = "<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq></Where></Query><RowLimit Paged=TRUE'>1</RowLimit></View>"
                                $spListItem = $spList.GetItems($Query)
                                [xml]$xmlProperty = $spListItem.Xml
                                if ($xmlProperty.xml.data.row.ows_Editor)
                                    {
                                    $ListLastItemModifiedBy = $xmlProperty.xml.data.row.ows_Editor.Split("#")[1]
                                    }
                                else
                                    {
                                    $ListLastItemModifiedBy = "error"
                                    }

                                }
                            else
                                {
                                $ListLastItemModifiedBy = "empty"
                                }
                            }
                        else
                            {
                            $ListLastItemModifiedBy = "not queried"
                            }
                        #region - query to get "last modified by" info
                        $spListChangeLogObject = New-Object -TypeName psobject
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name WebApplicationUrl -Value $spWebApplication.Url
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name SiteId -Value $spWeb.Site.ID.Guid
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $spWeb.Site.Url
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name SiteLastContentModifiedDate -Value $spSite.LastContentModifiedDate.ToString("yyyy-MM-ddThh:mm:ss")
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name WebId -Value $spWeb.ID.Guid
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $spWeb.Url
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name WebLastItemModifiedDate -Value $spWeb.LastItemModifiedDate.ToString("yyyy-MM-ddThh:mm:ss")
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name ListId -Value $spList.ID.Guid
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name ListUrl -Value ($spList.ParentWeb.Url.TrimEnd("/") + "/" + $spList.RootFolder.Url)
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name ListType -Value $spList.BaseTemplate
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name ListTitle -Value $spList.Title
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name ListItemCount -Value $ListItemCount
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name ListLastItemModifiedDate -Value $spList.LastItemModifiedDate.ToString("yyyy-MM-ddThh:mm:ss")
                        $spListChangeLogObject | Add-Member -MemberType NoteProperty -Name ListLastItemModifiedBy -Value $ListLastItemModifiedBy
                        $spObjects += $spListChangeLogObject
                        }
                    }
                $spWeb.Dispose()
                }
            $spSite.Dispose()
            }
        }
    else
        {
        if ($outputEnabled) {Write-Host ("NOTE: Skipping web application, no sites detected: " + $spWebApplication.Url) -foreGroundcolor Cyan}
        }
    }
$spObjects | Export-Csv -Path $reportFullPathName -Encoding UTF8 -NoTypeInformation
if ($outputEnabled) {Write-Host ("Output can be found here: " + $reportFullPathName)}
#endregion - Main
###########################################################################################################################################
}
#endregion - TRY
###########################################################################################################################################
#region - Catch & Finally
#region - Catch
catch
{
    if ($outputEnabled) {
    $Error[0]
    }
}
#endregion - Catch
#region - Finally
finally
{
    $ErrorActionPreference = "Continue";
    $endTime = (Get-Date)
    $timeSpan = ($($endTime - $startTime).TotalSeconds)
    if ($outputEnabled) {
    Write-Host ("Script Ended: ") -NoNewline -ForegroundColor Cyan
    Write-Host $endTime -ForegroundColor DarkYellow -NoNewline
    Write-Host (" (and took: ") -ForegroundColor Gray -NoNewline
    Write-Host ([math]::Round($timeSpan,3)) -NoNewline
    Write-Host (" Seconds)") -ForegroundColor Gray
    }
    if ($EnableLogging) {Stop-Transcript}
}
#endregion - Finally
#endregion - Catch & Finally
###########################################################################################################################################

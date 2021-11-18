<#
.NOTES
    #######################################################################################################################################
    # Author: Jim B.
    # Latest version can be found on Github: https://github.com/jbienstman/SMAT
    #######################################################################################################################################
    # Revision(s)
    # 1.0.0 2021-10-14  Initial Version
    # 1.1.0 2021-11-08  Including Get-spWebRoleAssignments Function inside script & several updates
    # 1.2.0 2021-11-08  Now checking whether the $reportPath exists (added Function: AskYesNoQuestion)
    # 1.2.1 2021-11-18  Updated Header Synopsis, Description, Example ...
    #######################################################################################################################################
.SYNOPSIS
    On a SharePoint Server 2010+ Farm (On-Premise), this script will iterate through all webs in all sites in all web applications and output a
    CSV report with the web-level "Role Assignments". A custom function "Get-spWebRoleAssignments" is used and included in this script.

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
    ReportPermissionAllWebsInFarm -reportPath "C:\TEMP\" -outputEnabled:$true
#>
###########################################################################################################################################
Param (
    [Parameter(Mandatory=$false, HelpMessage = "Folder location where report should be saved")][string]$reportPath = "C:\TEMP\" ,
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
Function Get-spWebRoleAssignments {
    <#
    .SYNOPSIS
        This function will retrieve the permissions on the SPWeb if not inherited
    .EXAMPLE
        Get-RoleAssignments -spWeb $spWeb
    .INPUTS
        SPWeb Object
    .OUTPUTS
        string
    .NOTES
        Author:  Jim B.
        Website: https://github.com/jbienstman
    #>
Param (
    [Parameter(Mandatory = $true)][object]$spWeb,
    [Parameter(Mandatory = $false)][bool]$showOutput = $false
    )
$roleAssignmentObjectArray = @()
if ($spWeb.HasUniquePerm -eq $true)
    {
    foreach ($RoleAssignment in $spWeb.RoleAssignments)
        {
        foreach ($RoleDefinitionBinding in $RoleAssignment.RoleDefinitionBindings)
            {
            if ($RoleAssignment.Member.IsDomainGroup -eq $false -and $RoleAssignment.Member.GetType().Name -eq "SPUser")
                {
                if ($showOutput) {Write-Host ("[" + $RoleAssignment.Member.GetType().Name + "] - " + $RoleAssignment.Member.DisplayName + " has " + $RoleDefinitionBinding.Name.ToString())}
                #
                $roleAssignmentObject = New-Object -TypeName psobject
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteId -Value $spWeb.Site.ID
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $spWeb.Site.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebId -Value $spWeb.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $spWeb.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name IsRootWeb -Value $spWeb.IsRootWeb
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name HasUniquePerm -Value $spWeb.HasUniquePerm
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Type -Value $RoleAssignment.Member.GetType().Name
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Member -Value $RoleAssignment.Member.DisplayName
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name RoleDefinitionBinding -Value $RoleDefinitionBinding.Name.ToString()
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name GroupMembers -Value "empty"
                }
            elseif ($RoleAssignment.Member.IsDomainGroup -eq $true -and $RoleAssignment.Member.GetType().Name -eq "SPUser")
                {
                if ($showOutput) {Write-Host ("[DomainGroup] - " + $RoleAssignment.Member.DisplayName + " has " + $RoleDefinitionBinding.Name.ToString())}
                #
                $roleAssignmentObject = New-Object -TypeName psobject
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteId -Value $spWeb.Site.ID
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $spWeb.Site.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebId -Value $spWeb.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $spWeb.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name IsRootWeb -Value $spWeb.IsRootWeb
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name HasUniquePerm -Value $spWeb.HasUniquePerm
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Type -Value "DomainGroup"
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Member -Value $RoleAssignment.Member.DisplayName
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name RoleDefinitionBinding -Value $RoleDefinitionBinding.Name.ToString()
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name GroupMembers -Value "empty"
                }
            else
                {
                if ($showOutput) {Write-Host ("[" + $RoleAssignment.Member.GetType().Name + "] - " + $RoleAssignment.Member.Name + " has " + $RoleDefinitionBinding.Name.ToString())}
                [string]$spGroupMembers = ""
                $spWeb.Groups[$RoleAssignment.Member.Name].Users | Select-Object -ExpandProperty DisplayName | ForEach-Object {$spGroupMembers += ($_ + ";")}
                #
                $roleAssignmentObject = New-Object -TypeName psobject
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteId -Value $spWeb.Site.ID
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $spWeb.Site.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebId -Value $spWeb.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $spWeb.Url
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name IsRootWeb -Value $spWeb.IsRootWeb
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name HasUniquePerm -Value $spWeb.HasUniquePerm
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Type -Value $RoleAssignment.Member.GetType().Name
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Member -Value $RoleAssignment.Member.Name
                $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name RoleDefinitionBinding -Value $RoleDefinitionBinding.Name.ToString()
                if ($spGroupMembers.Length -eq 0)
                    {
                    $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name GroupMembers -Value "empty"
                    }
                else
                    {
                    $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name GroupMembers -Value $spGroupMembers.TrimEnd(";")
                    }

                }
            $roleAssignmentObjectArray += $roleAssignmentObject
            }
        }
    }
else
    {
        $roleAssignmentObject = New-Object -TypeName psobject
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteId -Value $spWeb.Site.ID
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $spWeb.Site.Url
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebId -Value $spWeb.Url
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $spWeb.Url
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name IsRootWeb -Value $spWeb.IsRootWeb
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name HasUniquePerm -Value $spWeb.HasUniquePerm
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Type -Value "inherited"
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name Member -Value "inherited"
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name RoleDefinitionBinding -Value "inherited"
        $roleAssignmentObject | Add-Member -MemberType NoteProperty -Name GroupMembers -Value "inherited"
        $roleAssignmentObjectArray += $roleAssignmentObject
    }
return $roleAssignmentObjectArray
}
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
$spWebApplications = Get-SPWebApplication # -Identity $spWebApplicationUrl
if ($outputEnabled) {
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
$reportFullPathName = ($reportPath.Trim("\") + "\" + "FarmSitePermissions_" + (Get-Date -Format yyyyMMddHHmmss) + ".csv")
$spFarmRoleAssignmentObjects = @()
$i = 0
foreach ($spWebApplication in $spWebApplications)
    {
    $i++
    $spWebApplicationRoleAssignmentObjects = @()
    $spSites = $spWebApplication.Sites
    if ($outputEnabled) {Write-Host ("[" + $i + "/" + $spWebApplications.Count +  "] (web application) - `"" + $spWebApplication.Url + "`"") -foreGroundcolor Cyan}
    $j = 0
    foreach ($spSite in $spSites) {
        $j++
        if ($outputEnabled) {Write-Host (" - [" + $j + "/" + $spSites.Count.ToString() + "] (spSite) - `"" + $spSite.ServerRelativeUrl + "`"") -foreGroundcolor Yellow}
        #$spSite = Get-SPSite -Identity $spSite.Url
        $k = 0
        foreach ($spWeb in $spSite.AllWebs) {
            $k++
            if ($outputEnabled) {Write-Host (" - + [" + $k + "/" + $spSite.AllWebs.Count.ToString() + "] (spWeb) - `"" + $spweb.ServerRelativeUrl + "`"") -foreGroundcolor DarkGray}
            $spWebRoleAssignmentObject = Get-spWebRoleAssignments -spWeb $spWeb
            $spWebApplicationRoleAssignmentObjects += $spWebRoleAssignmentObject
            $spWeb.Dispose()
        }
        $spSite.Dispose()
    }
    if ($outputEnabled) {Write-Host (" Adding Web Application Role Assignments to Farm Report") -ForegroundColor DarkGray}
    $spFarmRoleAssignmentObjects += $spWebApplicationRoleAssignmentObjects
    $spWebApplication = $null
    $spWebApplicationRoleAssignmentObjects = $null
}
#region - report output
if ($outputEnabled) {Write-Host ("Writing to log: " + $reportFullPathName) -NoNewline}
$spFarmRoleAssignmentObjects | Export-Csv -Path $reportFullPathName -Encoding UTF8 -NoTypeInformation -NoClobber
$spFarmRoleAssignmentObjects = $null
[GC]::Collect()
if ($outputEnabled) {Write-Host ("...Done") -ForegroundColor Green}
if ($outputEnabled) {Write-Host ("Report can be found here: " + $reportFullPathName) -ForegroundColor Yellow}
#endregion - report output
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

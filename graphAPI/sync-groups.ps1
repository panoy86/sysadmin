<#
.DESCRIPTION
    Purpose of this script is to do a one-way sync from one to multiple groups onto a single group.
.PARAMETER SourceGroups
    Comma-separated list of groups to sync from, members will be in the target group, required.
.PARAMETER TargetGroup
    Group to sync to, required.
.PARAMETER ExceptGroups
    Comma-separated list of groups to exclude from the sync, members will be remove from target group, optional.
.PARAMETER EmailNotifier
    Email address to notify when there is an error, optional (not yet implemented).
.NOTES
    This requires an existing/authenticated Microsoft Graph PowerShell session with the appropriate permissions.
    App permissions required: Group.ReadWrite.All, Directory.Read.All
    EOL permissions required: (minimum) Recipient Management
    If target group in an Exchange object, we need to use Exchange Online Management shell to modify the membership.
    Modules required: Microsoft.Graph, ExchangeOnlineManagement
    Notes: with unified groups, if a member is an owner, this script cannot remove it as a member. Manually remove the owner first.
.EXAMPLE
    .\sync-groups.ps1 -SourceGroups "group1,group2" -TargetGroup "group3" -ExceptGroups "group4" -EmailNotifier "admin@contoso.com"
#>
#-- Example from Microsoft on how to create a service principal for Graph and EOL use:
#-- https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-service-principal-portal
#-- https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps

param (
    [string] $SourceGroups,  #-- Comma-separated list of groups to sync from
    [string] $TargetGroup,   #-- Group to sync to
    [string] $ExceptGroups,  #-- Comma-separated list of groups to exclude from the sync, will ignore if empty
    [string] $EmailNotifier  #-- Email address to notify when there is an error, will ignore if empty (not yet implemented)
)

$Script:UseExistingSession = $true  #-- Set to $false if you want the script to create a new session, or if you want to use this in a scheduled task with a stored credential
$Script:TempListGroups = @()   #-- Shared list of groups called by multiple functions
$Script:TempListMembers = @()  #-- Shared list of members of a group
$Script:TempHashMembers = @{}  #-- Shared hash table of members to avoid duplicates
$Script:TempHashGroups = @{}   #-- Sahred hash table of groups to avoid endless loops

#------------------------------------------------------------------------------
#-- Authenticate to Microsoft Graph and Exchange Online
#------------------------------------------------------------------------------
function AuthenticateToGraphAndEOL
{
    #-- Disconnect from any existing sessions
    Disconnect-MgGraph -ea SilentlyContinue

    #-- Change these to your service principal details
    $tenantId = "your tenant id"
    $appId = "your app id"
    $thumbprint = "your cert thumbprint"

    Connect-ExchangeOnline -AppId $appId -CertificateThumbprint $thumbprint -Organization "somedomain.onmicrosoft.com" -ShowBanner:$false -ErrorAction Stop
    Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -NoWelcome
    (Get-MgContext).Scopes | Sort-Object
}

#------------------------------------------------------------------------------
#-- Search the groups and confirm they exist
#------------------------------------------------------------------------------
function VerifyGroupList
{
    $returnValue = $true
    #-- Search for the source and target groups
    foreach ($entry in $Script:TempListGroups)
    {
        #-- Try to search via Alias/MailNickName, DisplayName, and Mail properties
        $groupName = $entry.Identity
        $group = $null
        $group = Get-MgGroup -ConsistencyLevel eventual -CountVariable groupCount -ErrorAction SilentlyContinue -Select Id,DisplayName,OnPremisesSyncEnabled -Filter "MailNickName eq '$groupName'"
        if ($null -eq $group) {$group = Get-MgGroup -ConsistencyLevel eventual -CountVariable groupCount -ErrorAction SilentlyContinue -Select Id,DisplayName,OnPremisesSyncEnabled -Filter "Mail eq '$groupName'"}
        if ($null -eq $group) {$group = Get-MgGroup -ConsistencyLevel eventual -CountVariable groupCount -ErrorAction SilentlyContinue -Select Id,DisplayName,OnPremisesSyncEnabled -Filter "DisplayName eq '$groupName'"}
        if ($null -ne $group -and $global:groupCount -eq 1)
        {            
            if ($group.Count -gt 1) {$returnValue = $false}
            $entry.Guid = $group.Id
            $entry.Cloud = (-not $group.OnPremisesSyncEnabled)
        }
        else
        {
            #-- Test if this is a dynamic distribution list in Exchange Online
            $group = Get-DynamicDistributionGroup -Identity $groupName -ErrorAction SilentlyContinue
            $group
            if ($null -ne $group)
            {
                $entry.Guid = $group.Guid.ToString()
                $entry.Cloud = $true
                $entry.DDL = $true
            } else {$returnValue = $false}
        }
    }
    return $returnValue
}

#------------------------------------------------------------------------------
#-- Get all the members of a group, recursively if you need to
#------------------------------------------------------------------------------
function GetMembersOfGroup
{
    param (
        [string] $GroupId  #-- Group to get members from
    )

    #-- Get the members at this level
    $members = Get-MgGroupMember -GroupId $GroupId -ConsistencyLevel eventual -All #-ErrorAction SilentlyContinue

    #-- Process the members of the group
    if ($null -ne $members)
    {
        foreach ($member in $members)
        {
            #-- Found a user account, change this if you want to include other object types
            if ($member['@odata.type'] -eq "#microsoft.graph.user")
            {
                #-- Only add to our list, if not already present
                $key = $member.Id.ToString()
                if (-not $Script:TempHashMembers.ContainsKey($key))
                {
                    $Script:TempListMembers += $key
                    $Script:TempHashMembers.Add($key, 1)
                }
            }
            #-- Recursively get members of groups, but only if not already processed
            if ($member['@odata.type'] -eq "#microsoft.graph.group" -and (-not $Script:TempHashGroups.ContainsKey($member.Id)))
            {
                #Write-Host "$($member.AdditionalProperties['displayName']) " -NoNewline
                #$member.AdditionalProperties | fl
                $Script:TempHashGroups.Add($member.Id, 1)
                GetMembersOfGroup -GroupId $member.Id
            }
        }
    }
}

#------------------------------------------------------------------------------
#-- Get all the members of a dynamic DL
#------------------------------------------------------------------------------
function GetMembersOfDDL
{
    param (
        [string] $DDLIdentity  #-- Dynamic DL to get members from
    )

    #-- Get the members at this level
    $members = Get-Recipient -RecipientPreviewFilter (Get-DynamicDistributionGroup -Identity $DDLIdentity).RecipientFilter -ResultSize Unlimited #-ErrorAction SilentlyContinue

    #-- Process the members of the group
    if ($null -ne $members)
    {
        foreach ($member in $members)
        {
            #-- Found a user account, change this if you want to include other object types
            if ($member.RecipientType -eq "UserMailbox" -or $member.RecipientType -eq "MailUser")
            {
                #-- Only add to our list, if not already present
                $key = $member.ExternalDirectoryObjectId.ToString()
                if (-not $Script:TempHashMembers.ContainsKey($key))
                {
                    $Script:TempListMembers += $key
                    $Script:TempHashMembers.Add($key, 1)
                }
            }
        }
    }
}

#------------------------------------------------------------------------------
#-- Get total and unique members of a set of groups
#------------------------------------------------------------------------------
function GetAllMembersOfGroups
{
    param (
        [string] $direction  #-- Direction of the groups to process: source, target, except
    )

    #-- Reset the temporary variables we need to maintain unique members and prevent loops
    $Script:TempListMembers = @()
    $Script:TempHashMembers = @{}
    $Script:TempHashGroups = @{}

    #-- Loop thru the group IDs and get the members
    foreach ($group in $Script:TempListGroups | Where-Object {$_.Direction -eq $direction})
    {
        $Script:TempHashGroups.Add($group.Guid, 1)
        #Write-Host "$($group.Identity) " -ForegroundColor Cyan -NoNewline
        if (-not $group.DDL) {GetMembersOfGroup -GroupId $group.Guid}
        else {GetMembersOfDDL -DDLIdentity $group.Identity}
    }
}

#------------------------------------------------------------------------------
#-- Main program
#------------------------------------------------------------------------------
if ([int]$PSVersionTable.PSVersion.Major -ge 7)
{
    Import-Module PSReadLine -Force  #-- Fixes progress bar issues in PowerShell 7+
    $PSStyle.Progress.View = "Minimal"  #-- Other value: "Minimal", only works in PowerShell 7.2+
}
$ProgressPreference = "Continue"
$dateStart = (Get-Date)

#-- Normalize the parameters
$rTmpGroups = @()
$SourceGroups -split ' ' | ForEach-Object {$rTmpGroups += $_.Trim()}
$TargetGroup = $TargetGroup.Trim()

#-- Show command line usage
if ($SourceGroups.Trim().Length -eq 0 -or $rTmpGroups.Count -eq 0 -or $TargetGroup.Length -eq 0)
{
    Write-Host ' '
    Write-Host "Usage: " -NoNewline
    Write-Host ".\sync-groups.ps1"
    Write-Host "      -SourceGroups " -NoNewline -ForegroundColor Yellow
    Write-Host "group1,group2,etc (members will be in the target group, required)" 
    Write-Host "      -TargetGroup " -NoNewline -ForegroundColor Yellow
    Write-Host "group-to-sync-to (target group, required)"
    Write-Host "      -ExceptGroups " -NoNewline -ForegroundColor Yellow
    Write-Host "group7,group8,etc (members will be remove from target group, optional)"
    Write-Host "      -EmailNotifier " -NoNewline -ForegroundColor Yellow
    Write-Host "email-address-to-notify (in case of errors, optional)"
    Write-Host ' '
    return
}

#-- Connect to Microsoft Graph and Exchange Online Management shell
if (-not $Script:UseExistingSession)
{
    AuthenticateToGraphAndEOL
}

#-- Prepare our list of groups, combine the sources and target into a single list
$Script:TempListGroups = @()
$rTmpGroups | ForEach-Object{
    $Script:TempListGroups += [PSCustomObject] @{
        Identity = $_.Trim()
        Direction = "source"
        Guid = ''
        Cloud = ''
        Type = ''
        DDL = $false #-- Dynamic Distribution List, special case
    }
}
$Script:TempListGroups += [PSCustomObject] @{
    Identity = $TargetGroup
    Direction = "target"
    Guid = ''
    Cloud = ''
    Type = ''
}

#-- Check if the optional parameters are provided
if ($ExceptGroups -and $ExceptGroups.Trim().Length -gt 0)
{
    $rTmpGroups = @()
    $ExceptGroups -split ' ' | ForEach-Object {$rTmpGroups += $_.Trim()}
    $rTmpGroups | ForEach-Object{
        $Script:TempListGroups += [PSCustomObject] @{
            Identity = $_
            Direction = "except"
            Guid = ''
            Cloud = ''
            Type = ''
        }
    }
}

#-- Verify if all groups are valid
if (-not (VerifyGroupList))
{
    Write-Host "One or more groups do not exist, or an entry match multiple groups." -ForegroundColor Red
    Write-Host "Please check the group names." -ForegroundColor Red
    $Script:TempListGroups
    return
}
$Script:TempListGroups | Format-Table -AutoSize -Wrap

#-- Loop thru our list our source groups
GetAllMembersOfGroups "source"
Write-Host "source count:" $Script:TempListMembers.Count
$rSourceMembers = $Script:TempListMembers

#-- Add the hash of source groups to a running total of groups
$hashSourceAndExceptGroups = $Script:TempHashGroups.Clone()

#-- Loop thru our list of groups to exclude, if any
if ($ExceptGroups -and $ExceptGroups.Trim().Length -gt 0)
{
    GetAllMembersOfGroups "except"
    Write-Host "except count:" $Script:TempListMembers.Count
    $rExceptMembers = $Script:TempListMembers

    #-- Remove the members from the source list, if they are in the except list
    if ($rExceptMembers.Count -gt 0)
    {
        #-- Prepare the hash table of members to remove, for faster lookups
        $hashMembersToRemove = @{}
        $rExceptMembers | ForEach-Object {$hashMembersToRemove.Add($_, 1)}

        #-- Loop thru the source members and create a temp list without the removed-members
        $removeCount = 0
        $rTmp = @()
        foreach ($member in $rSourceMembers)
        {
            if (-not $hashMembersToRemove.ContainsKey($member)) {$rTmp += $member}
            else {$removeCount++}
        }
    }
    #-- Replace the source members with the filtered list   
    $rSourceMembers = $rTmp
    Write-Host "removed: $($removeCount), new count is $($rSourceMembers.Count)"
}

#-- Another check, return an error if the target group is in any of the source and except groups (even child groups)
#-- Add the results of the except groups to the hash of source groups first, then check
foreach ($hashItem in $Script:TempHashGroups.GetEnumerator())
{
    if (-not $hashSourceAndExceptGroups.ContainsKey($hashItem.Key))
    {
        $hashSourceAndExceptGroups.Add($hashItem.Key, 1)
    }
}
if ($hashSourceAndExceptGroups.Contains(($Script:TempListGroups | Where-Object {$_.Direction -eq "target"}).Guid))
{
    Write-Host "Target group is in the source or except groups, please remove it from the list." -ForegroundColor Red
    return
}

#-- Get the current members of the target group
GetAllMembersOfGroups "target"
Write-Host "target count:" $Script:TempListMembers.Count
$targetMembers = $Script:TempListMembers

#-- Here is where it gets interesting... if the target group is mail-enabled, we cannot modify the membership
#-- using Microsoft Graph, instead we have to use Exchange Online Management shell.
#-- I wish Microsoft can make this more consistent...
$targetGuid = ($Script:TempListGroups | Where-Object {$_.Direction -eq "target"}).Guid
$targetGroupType = ''
$targetGroup = $null
$targetGroup = Get-DistributionGroup -Identity $targetGuid -ErrorAction SilentlyContinue
if ($null -eq $targetGroup) {$targetGroup = Get-UnifiedGroup -Identity $targetGuid -ErrorAction SilentlyContinue}
if ($null -ne $targetGroup)
{
    # It's a mail-enabled group.
    if ($targetGroup.RecipientTypeDetails -match "DistributionGroup")
    {
        $targetGroupType = "DistributionGroup"
    }
    elseif ($targetGroup.RecipientTypeDetails -match "GroupMailbox")
    {
        $targetGroupType = "UnifiedGroup"
    }
    # It's not a group we can modify or taken into account, throw an error
    else
    {
        Write-Host "Target group is not a known type, $($targetGroup.RecipientTypeDetails)" -ForegroundColor Red
        return
    }
}
# Let's identity this as a Microsoft Graph-capable group
else
{
    $targetGroupType = "MicrosoftGraphGroup"
}
Write-Host "target group type: $targetGroupType" -ForegroundColor Green

# Now that we have identified the target group type, remove the members that are not in the source list
$removeCount = 0
$hashSourceMembers = @{}
$rSourceMembers | ForEach-Object {$hashSourceMembers.Add($_, 1)}
$nCtr = 0
foreach ($member in $targetMembers)
{
    # Show progress
    $nCtr++
    Write-Progress -Activity "Removing members from target group" -Status "Processing member $nCtr of $($targetMembers.Count)" -PercentComplete (($nCtr / $targetMembers.Count) * 100)

    # If the member is already in the source list, skip it
    if (-not $hashSourceMembers.ContainsKey($member))
    {
        # Remove the member from the target group
        if ($targetGroupType -eq "DistributionGroup")
        {
            # Use Exchange Online Management shell to remove the member
            Remove-DistributionGroupMember -Identity $targetGuid -Member $member -Confirm:$false -BypassSecurityGroupManagerCheck #-ErrorAction SilentlyContinue
            $removeCount++
        }
        elseif ($targetGroupType -eq "UnifiedGroup")
        {
            # Use Exchange Online Management shell to remove the member
            Remove-UnifiedGroupLinks -Identity $targetGuid -LinkType Members -Links $member -Confirm:$false #-ErrorAction SilentlyContinue
            $removeCount++
        }
        else
        {
            # Use Microsoft Graph to remove the member
            Remove-MgGroupMemberByRef -GroupId $targetGuid -DirectoryObjectId $member #-ErrorAction SilentlyContinue
            $removeCount++
        }
    }
}
Write-Progress -Activity "Removing members from target group" -Completed -Status "Processing complete"
Write-Host "removed: $($removeCount) from target group"

#-- Get an update list of target members after the removals
#-- Get the current members of the target group
GetAllMembersOfGroups "target"
$targetMembers = $Script:TempListMembers

#-- Final step, add the members from the source list to the target group
$addCount = 0
$hashTargetMembers = @{}
$targetMembers | ForEach-Object {$hashTargetMembers.Add($_, 1)}
$nCtr = 0
foreach ($member in $rSourceMembers)
{
    # Show progress
    $nCtr++
    Write-Progress -Activity "Adding members to target group" -Status "Processing member $nCtr of $($rSourceMembers.Count)" -PercentComplete (($nCtr / $rSourceMembers.Count) * 100)

    # If the member is already in the target group, skip it
    if (-not $hashTargetMembers.Contains($member))
    {
        # Add the member to the target group
        if ($targetGroupType -eq "DistributionGroup")
        {
            # Use Exchange Online Management shell to add the member
            Add-DistributionGroupMember -Identity $targetGuid -Member $member -BypassSecurityGroupManagerCheck #-ErrorAction SilentlyContinue
            $addCount++
        }
        elseif ($targetGroupType -eq "UnifiedGroup")
        {
            # Use Exchange Online Management shell to add the member
            Add-UnifiedGroupLinks -Identity $targetGuid -LinkType Members -Links $member #-ErrorAction SilentlyContinue
            $addCount++
        }
        else
        {
            # Use Microsoft Graph to add the member
            New-MgGroupMember -GroupId $targetGuid -DirectoryObjectId $member #-ErrorAction SilentlyContinue
            $addCount++
        }
    }
}
Write-Progress -Activity "Adding members to target group" -Completed -Status "Processing complete"
Write-Host "added: $($addCount) to target group"
Write-Host "runtime" ([int]((Get-Date) - $dateStart).TotalMinutes) "minutes, end."
#Disconnect-MgGraph -ea SilentlyContinue
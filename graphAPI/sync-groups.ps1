#-- Purpose of this script is to do a one-way sync from one to multiple groups onto a single group.
#-- It supports security groups, distribution lists, and Microsoft 365 groups.
#-- The destination group is only one level deep, so it does not support nested groups.
#-- The source groups can be nested.
#-- The destination group is cloud-only.
#-- Scope is cloud-only or cloud replicated groups, it cannot connect to on-premises AD-DS.
#-- Scope is user accounts only, it will not sync contacts or other objects (easy to change though).
#-- Sample use case: use several DLs and security groups to create a single set of members for a Teams group.
#-- App permissions required: Group.ReadWrite.All, Directory.Read.All
#-- EOL permissions required: (minimum) Recipient Management
#-- If target group in an Exchange object, we need to use Exchange Online Management shell to modify the membership.
#-- Modules required: Microsoft.Graph, ExchangeOnlineManagement
#-- Notes: with unified groups, if a member is an owner, this script cannot remove it as a member. Manually remove the owner first.

param (
    [string] $SourceGroups,  #-- Comma-separated list of groups to sync from
    [string] $TargetGroup,   #-- Group to sync to
    [string] $ExceptGroups,  #-- Comma-separated list of groups to exclude from the sync, will ignore if empty
    [string] $EmailNotifier  #-- Email address to notify when there is an error, will ignore if empty (not yet implemented)
)

$script:tempListGroups = @()   #-- Shared list of groups called by multiple functions
$script:tempListMembers = @()  #-- Shared list of members of a group
$script:tempHashMembers = @{}  #-- Shared hash table of members to avoid duplicates
$script:tempHashGroups = @{}   #-- Sahred hash table of groups to avoid endless loops

#------------------------------------------------------------------------------
#-- Search the groups and confirm they exist
#------------------------------------------------------------------------------
function VerifyGroupList
{
    $returnValue = $true
    #-- Search for the source and target groups
    foreach ($entry in $script:tempListGroups)
    {
        #-- Try to search via Alias/MailNickName, DisplayName, and Mail properties
        $sTmp = $entry.Identity
        $group = Get-MgGroup -ConsistencyLevel eventual -CountVariable groupCount -Search "DisplayName:$sTmp" -Select Id,DisplayName,OnPremisesSyncEnabled -ErrorAction SilentlyContinue
        if ($null -eq $group) {$group = Get-MgGroup -ConsistencyLevel eventual -CountVariable groupCount -Search "MailNickName:$sTmp" -ErrorAction SilentlyContinue}
        if ($null -eq $group) {$group = Get-MgGroup -ConsistencyLevel eventual -CountVariable groupCount -Search "Mail:$sTmp" -ErrorAction SilentlyContinue}
        if ($null -ne $group)
        {
            if ($group.Count -gt 1) {$returnValue = $false}
            $entry.Guid = $group.Id
            $entry.Cloud = (-not $group.OnPremisesSyncEnabled)
        }
        else {$returnValue = $false}
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
    $rMembers = Get-MgGroupMember -GroupId $GroupId -ConsistencyLevel eventual -All #-ErrorAction SilentlyContinue

    #-- Process the members of the group
    if ($null -ne $rMembers)
    {
        foreach ($member in $rMembers)
        {
            #-- Found a user account, change this if you want to include other object types
            if ($member['@odata.type'] -eq "#microsoft.graph.user")
            {
                #-- Only add to our list, if not already present
                $key = $member.Id.ToString()
                if (-not $script:tempHashMembers.ContainsKey($key))
                {
                    $script:tempListMembers += $key
                    $script:tempHashMembers.Add($key, 1)
                }
            }
            #-- Recursively get members of groups, but only if not already processed
            if ($member['@odata.type'] -eq "#microsoft.graph.group" -and (-not $script:tempHashGroups.ContainsKey($member.Id)))
            {
                Write-Host "$($member.AdditionalProperties['displayName']) " -NoNewline
                #$member.AdditionalProperties | fl
                $script:tempHashGroups.Add($member.Id, 1)
                GetMembersOfGroup -GroupId $member.Id
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
    $script:tempListMembers = @()
    $script:tempHashMembers = @{}
    $script:tempHashGroups = @{}

    #-- Loop thru the group IDs and get the members
    foreach ($group in $script:tempListGroups | Where-Object {$_.Direction -eq $direction})
    {
        $script:tempHashGroups.Add($group.Guid, 1)
        Write-Host "$($group.Identity) " -ForegroundColor Cyan -NoNewline
        GetMembersOfGroup -GroupId $group.Guid
    }
}

#------------------------------------------------------------------------------
#-- Main program
#------------------------------------------------------------------------------
$PSStyle.Progress.View = "Classic"  #-- Other value: "Minimal", only works in PowerShell 7.2+
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

#-- Prepare our list of groups, combine the sources and target into a single list
$script:tempListGroups = @()
$rTmpGroups | ForEach-Object{
    $script:tempListGroups += [PSCustomObject] @{
        Identity = $_.Trim()
        Direction = "source"
        Guid = ''
        Cloud = ''
    }
}
$script:tempListGroups += [PSCustomObject] @{
    Identity = $TargetGroup
    Direction = "target"
    Guid = ''
    Cloud = ''
}

#-- Check if the optional parameters are provided
if ($ExceptGroups -and $ExceptGroups.Trim().Length -gt 0)
{
    $rTmpGroups = @()
    $ExceptGroups -split ' ' | ForEach-Object {$rTmpGroups += $_.Trim()}
    $rTmpGroups | ForEach-Object{
        $script:tempListGroups += [PSCustomObject] @{
            Identity = $_
            Direction = "except"
            Guid = ''
            Cloud = ''
        }
    }
}

#-- Verify if all groups are valid
if (-not (VerifyGroupList))
{
    Write-Host "One or more groups do not exist, or an entry match multiple groups." -ForegroundColor Red
    Write-Host "Please check the group names." -ForegroundColor Red
    $script:tempListGroups
    return
}
$script:tempListGroups | Format-Table -AutoSize -Wrap

#-- Loop thru our list our source groups
GetAllMembersOfGroups "source"
Write-Host ' '
Write-Host "source count:" $script:tempListMembers.Count
$rSourceMembers = $script:tempListMembers

#-- Add the hash of source groups to a running total of groups
$hashSourceAndExceptGroups = $script:tempHashGroups.Clone()

#-- Loop thru our list of groups to exclude, if any
if ($ExceptGroups -and $ExceptGroups.Trim().Length -gt 0)
{
    GetAllMembersOfGroups "except"
    Write-Host ' '
    Write-Host "except count:" $script:tempListMembers.Count
    $rExceptMembers = $script:tempListMembers

    #-- Remove the members from the source list, if they are in the except list
    if ($rExceptMembers.Count -gt 0)
    {
        #-- Prepare the hash table of members to remove, for faster lookups
        $hashMembersToRemove = @{}
        $rExceptMembers | ForEach-Object {$hashMembersToRemove.Add($_, 1)}

        #-- Loop thru the source members and create a temp list without the removed-members
        $intRemoveCount = 0
        $rTmp = @()
        foreach ($member in $rSourceMembers)
        {
            if (-not $hashMembersToRemove.ContainsKey($member)) {$rTmp += $member}
            else {$intRemoveCount++}
        }
    }
    #-- Replace the source members with the filtered list   
    $rSourceMembers = $rTmp
    Write-Host "removed: $($intRemoveCount), new count is $($rSourceMembers.Count)"
}

#-- Another check, return an error if the target group is in any of the source and except groups (even child groups)
#-- Add the results of the except groups to the hash of source groups first, then check
foreach ($hashItem in $script:tempHashGroups.GetEnumerator())
{
    if (-not $hashSourceAndExceptGroups.ContainsKey($hashItem.Key))
    {
        $hashSourceAndExceptGroups.Add($hashItem.Key, 1)
    }
}
if ($hashSourceAndExceptGroups.Contains(($script:tempListGroups | Where-Object {$_.Direction -eq "target"}).Guid))
{
    Write-Host "Target group is in the source or except groups, please remove it from the list." -ForegroundColor Red
    return
}

#-- Get the current members of the target group
GetAllMembersOfGroups "target"
Write-Host ' '
Write-Host "target count:" $script:tempListMembers.Count
$rTargetMembers = $script:tempListMembers

#-- Here is where it gets interesting... if the target group is mail-enabled, we cannot modify the membership
#-- using Microsoft Graph, instead we have to use Exchange Online Management shell.
#-- I wish Microsoft can make this more consistent...
$strTargetGuid = ($script:tempListGroups | Where-Object {$_.Direction -eq "target"}).Guid
$strTargetGroupType = ''
$objTargetGroup = $null
$objTargetGroup = Get-DistributionGroup -Identity $strTargetGuid -ErrorAction SilentlyContinue
if ($null -eq $objTargetGroup) {$objTargetGroup = Get-UnifiedGroup -Identity $strTargetGuid -ErrorAction SilentlyContinue}
if ($null -ne $objTargetGroup)
{
    #-- It's a mail-enabled group.
    if ($objTargetGroup.RecipientTypeDetails -match "DistributionGroup")
    {
        $strTargetGroupType = "DistributionGroup"
    }
    elseif ($objTargetGroup.RecipientTypeDetails -match "GroupMailbox")
    {
        $strTargetGroupType = "UnifiedGroup"
    }
    else  #-- It's not a group we can modify or taken into account, throw an error
    {
        Write-Host "Target group is not a known type, $($objTargetGroup.RecipientTypeDetails)" -ForegroundColor Red
        return
    }
}
else #-- Let's identity this as a Microsoft Graph-capable group
{
    $strTargetGroupType = "MicrosoftGraphGroup"
}
Write-Host "target group type: $strTargetGroupType" -ForegroundColor Green

#-- Now that we have identified the target group type, remove the members that are not in the source list
$intRemoveCount = 0
$hashSourceMembers = @{}
$rSourceMembers | ForEach-Object {$hashSourceMembers.Add($_, 1)}
foreach ($member in $rTargetMembers)
{
    if (-not $hashSourceMembers.ContainsKey($member))
    {
        #-- Remove the member from the target group
        if ($strTargetGroupType -eq "DistributionGroup")
        {
            #-- Use Exchange Online Management shell to remove the member
            Remove-DistributionGroupMember -Identity $strTargetGuid -Member $member -Confirm:$false -BypassSecurityGroupManagerCheck #-ErrorAction SilentlyContinue
            $intRemoveCount++
        }
        elseif ($strTargetGroupType -eq "UnifiedGroup")
        {
            #-- Use Exchange Online Management shell to remove the member
            Remove-UnifiedGroupLinks -Identity $strTargetGuid -LinkType Members -Links $member -Confirm:$false #-ErrorAction SilentlyContinue
            $intRemoveCount++
        }
        else
        {
            #-- Use Microsoft Graph to remove the member
            Remove-MgGroupMemberByRef -GroupId $strTargetGuid -DirectoryObjectId $member #-ErrorAction SilentlyContinue
            $intRemoveCount++
        }
    }
}
Write-Host "removed: $($intRemoveCount) from target group"

#-- Get an update list of target members after the removals
#-- Get the current members of the target group
GetAllMembersOfGroups "target"
Write-Host ' '
Write-Host "new target count:" $script:tempListMembers.Count "(replication delays may cause initial incorrect results)"
$rTargetMembers = $script:tempListMembers

#-- Final step, add the members from the source list to the target group
$intAddCount = 0
$hashTargetMembers = @{}
$rTargetMembers | ForEach-Object {$hashTargetMembers.Add($_, 1)}
foreach ($member in $rSourceMembers)
{
    #-- If the member is already in the target group, skip it
    if (-not $hashTargetMembers.Contains($member))
    {
        #-- Add the member to the target group
        if ($strTargetGroupType -eq "DistributionGroup")
        {
            #-- Use Exchange Online Management shell to add the member
            Add-DistributionGroupMember -Identity $strTargetGuid -Member $member -BypassSecurityGroupManagerCheck #-ErrorAction SilentlyContinue
            $intAddCount++
        }
        elseif ($strTargetGroupType -eq "UnifiedGroup")
        {
            #-- Use Exchange Online Management shell to add the member
            Add-UnifiedGroupLinks -Identity $strTargetGuid -LinkType Members -Links $member #-ErrorAction SilentlyContinue
            $intAddCount++
        }
        else
        {
            #-- Use Microsoft Graph to add the member
            New-MgGroupMember -GroupId $strTargetGuid -DirectoryObjectId $member #-ErrorAction SilentlyContinue
            $intAddCount++
        }
    }
}
Write-Host "added: $($intAddCount) to target group"
Write-Host "runtime" ([int]((Get-Date) - $dateStart).TotalMinutes) "minutes, end."

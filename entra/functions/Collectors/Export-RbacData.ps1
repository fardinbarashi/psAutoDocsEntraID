# Autodoc v1.0
function Export-RbacData {
    <#
        Section 1.6 - RBAC / PIM: active and eligible directory role assignments
        plus role-assignable groups, with principals resolved to names/types.
        Requires RoleManagement.Read.Directory, Directory.Read.All and AccessReview.Read.All.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths
    )

    $Section = 'Section 1.6 : RBAC / PIM'
    try {
        Write-Host "Start $Section" -ForegroundColor Yellow
        Write-Host "Build RBAC lookup tables..." -ForegroundColor Yellow

        $userLookup = @{}
        foreach ($user in (Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName")) {
            if (-not [string]::IsNullOrWhiteSpace([string]$user.Id)) {
                $userLookup[[string]$user.Id] = if ($user.UserPrincipalName) { "$($user.DisplayName) <$($user.UserPrincipalName)>" } else { $user.DisplayName }
            }
        }

        $groupsAll = Get-MgGroup -All -Property "Id,DisplayName,IsAssignableToRole,GroupTypes,MembershipRule"
        $groupLookup = @{}
        foreach ($group in $groupsAll) {
            if (-not [string]::IsNullOrWhiteSpace([string]$group.Id)) { $groupLookup[[string]$group.Id] = $group.DisplayName }
        }

        $spLookup = @{}
        foreach ($sp in (Get-MgServicePrincipal -All -Property "Id,DisplayName,AppId")) {
            if (-not [string]::IsNullOrWhiteSpace([string]$sp.Id)) { $spLookup[[string]$sp.Id] = $sp.DisplayName }
        }

        $principalLookups = @{ User = $userLookup; Group = $groupLookup; Sp = $spLookup }

        $roleDefinitionLookup = @{}
        try {
            foreach ($roleDef in (Get-MgRoleManagementDirectoryRoleDefinition -All)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$roleDef.Id)) { $roleDefinitionLookup[[string]$roleDef.Id] = $roleDef.DisplayName }
            }
        }
        catch {
            Write-Host "Could not read unified role definitions: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        Write-Host "Get role-assignable groups..." -ForegroundColor Yellow
        $roleAssignableGroups = $groupsAll | Where-Object { $_.IsAssignableToRole -eq $true }
        $roleAssignableGroupData = foreach ($group in $roleAssignableGroups) {
            $isDynamic = $null -ne $group.MembershipRule -and $group.MembershipRule.Trim().Length -gt 0
            [pscustomobject]@{
                RecordType           = "RoleAssignableGroup"
                RoleDefinitionName   = $null
                RoleDefinitionId     = $null
                PrincipalId          = $group.Id
                PrincipalDisplayName = $group.DisplayName
                PrincipalType        = "Group"
                AssignmentType       = "RoleAssignableGroup"
                MemberType           = if ($isDynamic) { "Dynamic" } else { "Assigned" }
                Status               = $null
                StartDateTime        = $null
                EndDateTime          = $null
                JITActivated         = $null
            }
        }

        Write-Host "Get active role assignment schedule instances..." -ForegroundColor Yellow
        $activeRoleAssignments = @()
        try { $activeRoleAssignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All }
        catch { Write-Host "Could not read active role assignments: $($_.Exception.Message)" -ForegroundColor Yellow }

        $activeRoleAssignmentData = foreach ($assignment in $activeRoleAssignments) {
            $principalId      = [string]$assignment.PrincipalId
            $roleDefinitionId = [string]$assignment.RoleDefinitionId
            $roleName = if ($roleDefinitionLookup.ContainsKey($roleDefinitionId)) { $roleDefinitionLookup[$roleDefinitionId] } else { $roleDefinitionId }

            $jitIndicator = $null
            if     ($assignment.AssignmentType -eq "Activated") { $jitIndicator = "Yes" }
            elseif ($assignment.AssignmentType)                 { $jitIndicator = "No" }

            [pscustomobject]@{
                RecordType           = "RoleAssignmentActive"
                RoleDefinitionName   = $roleName
                RoleDefinitionId     = $roleDefinitionId
                PrincipalId          = $principalId
                PrincipalDisplayName = Resolve-PrincipalName -PrincipalId $principalId -Lookups $principalLookups
                PrincipalType        = Resolve-PrincipalType -PrincipalId $principalId -Lookups $principalLookups
                AssignmentType       = $assignment.AssignmentType
                MemberType           = $assignment.MemberType
                Status               = $assignment.Status
                StartDateTime        = $assignment.StartDateTime
                EndDateTime          = $assignment.EndDateTime
                JITActivated         = $jitIndicator
            }
        }

        Write-Host "Get eligible role assignment schedule instances..." -ForegroundColor Yellow
        $eligibleRoleAssignments = @()
        try { $eligibleRoleAssignments = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All }
        catch { Write-Host "Could not read eligible role assignments: $($_.Exception.Message)" -ForegroundColor Yellow }

        $eligibleRoleAssignmentData = foreach ($assignment in $eligibleRoleAssignments) {
            $principalId      = [string]$assignment.PrincipalId
            $roleDefinitionId = [string]$assignment.RoleDefinitionId
            $roleName = if ($roleDefinitionLookup.ContainsKey($roleDefinitionId)) { $roleDefinitionLookup[$roleDefinitionId] } else { $roleDefinitionId }

            [pscustomobject]@{
                RecordType           = "RoleAssignmentEligible"
                RoleDefinitionName   = $roleName
                RoleDefinitionId     = $roleDefinitionId
                PrincipalId          = $principalId
                PrincipalDisplayName = Resolve-PrincipalName -PrincipalId $principalId -Lookups $principalLookups
                PrincipalType        = Resolve-PrincipalType -PrincipalId $principalId -Lookups $principalLookups
                AssignmentType       = "Eligible"
                MemberType           = $assignment.MemberType
                Status               = $assignment.Status
                StartDateTime        = $assignment.StartDateTime
                EndDateTime          = $assignment.EndDateTime
                JITActivated         = "No"
            }
        }

        $rbacCsv  = @()
        $rbacCsv += $activeRoleAssignmentData
        $rbacCsv += $eligibleRoleAssignmentData
        $rbacCsv += $roleAssignableGroupData
        $rbacCsv  = $rbacCsv | Sort-Object RoleDefinitionName, PrincipalDisplayName

        Write-Host "Export RBAC data to files..." -ForegroundColor Yellow
        Export-InventoryData -Data $rbacCsv -BaseName 'RBAC' -Paths $Paths

        Write-Host "Done. RBAC records: $(@($rbacCsv).Count)." -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green

        return @{ Rbac = $rbacCsv }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}

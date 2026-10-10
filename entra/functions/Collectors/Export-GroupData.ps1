# Autodoc v1.0
function Export-GroupData {
    <#
        Section 1.3 - Groups with owners, members, role assignments, licenses,
        welcome-email status and summary counts.
        Requires Group.Read.All, Directory.Read.All and RoleManagement.Read.Directory.

        -UserData is the output of Export-UserData; it seeds a lookup so member
        enrichment avoids re-querying users already collected.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths,
        [hashtable]$LicenseRef,
        [Parameter(Mandatory)][hashtable]$GraphRef,
        $UserData
    )

    $Section = 'Section 1.3 : Groups data'
    try {
        Write-Host "Start $Section" -ForegroundColor Yellow

        if (-not $LicenseRef) {
            $LicenseRef = @{ SkuByGuid = @{}; SkuByStringId = @{}; PlanById = @{}; PlanByName = @{} }
        }

        Write-Host "Get Groups data..."
        $entraGroups = Get-MgGroup -All -Property @(
            "id", "displayName", "description", "createdDateTime",
            "groupTypes", "mailEnabled", "securityEnabled", "mailNickname",
            "visibility", "membershipRule", "membershipRuleProcessingState",
            "isAssignableToRole", "resourceProvisioningOptions",
            "onPremisesSyncEnabled", "assignedLicenses"
        )

        Write-Host "Build group basic information..."
        $groupBasicInfo = [pscustomobject]@{
            TotalGroups      = $entraGroups.Count
            M365Groups       = ($entraGroups | Where-Object { $_.GroupTypes -contains "Unified" }).Count
            SecurityGroups   = ($entraGroups | Where-Object { $_.SecurityEnabled -eq $true }).Count
            DynamicGroups    = ($entraGroups | Where-Object { $null -ne $_.MembershipRule -and $_.MembershipRule.Trim().Length -gt 0 }).Count
            CloudGroups      = ($entraGroups | Where-Object { $_.OnPremisesSyncEnabled -ne $true }).Count
            OnPremisesGroups = ($entraGroups | Where-Object { $_.OnPremisesSyncEnabled -eq $true }).Count
        }

        Write-Host "Building user details lookup for group membership processing..." -ForegroundColor Yellow
        $userDetailsLookup = @{}
        foreach ($u in $UserData) {
            $userDetailsLookup[$u.Id] = [pscustomobject]@{
                Id                 = $u.Id
                DisplayName        = $u.DisplayName
                UserPrincipalName  = $u.UserPrincipalName
                Department         = $u.Department
                OfficeLocation     = $u.OfficeLocation
                Manager            = $u.Manager
                ManagerDisplayName = $u.ManagerDisplayName
            }
        }

        Write-Host "Build group data information..."
        $allGroupsCsv     = @()
        $allGroupMembers  = @()
        $groupWelcomeMail = @()

        foreach ($entraGroup in $entraGroups) {
            Write-Host "Process: $($entraGroup.DisplayName)"

            $isDynamic = $null -ne $entraGroup.MembershipRule -and $entraGroup.MembershipRule.Trim().Length -gt 0
            if     ($isDynamic)                                 { $groupCategory = "Dynamic" }
            elseif ($entraGroup.GroupTypes -contains "Unified") { $groupCategory = "M365" }
            elseif ($entraGroup.SecurityEnabled)                { $groupCategory = "Security" }
            else                                                { $groupCategory = "Other" }

            $source          = Get-GroupSource -OnPremisesSyncEnabled $entraGroup.OnPremisesSyncEnabled
            $owners          = Get-AllGroupOwners -GroupId $entraGroup.Id
            $memberOfInfo    = Get-GroupMemberOfInfo -GroupId $entraGroup.Id
            $roleAssignments = Get-GroupRoleAssignments -GroupId $entraGroup.Id
            $applications    = Get-GroupApplications -GroupId $entraGroup.Id
            $licenses        = Get-GroupLicenseNames -Group $entraGroup -LicenseRef $LicenseRef -GraphRef $GraphRef
            $ownerUpns       = ($owners | Where-Object { $_.UserPrincipalName } | Select-Object -ExpandProperty UserPrincipalName) -join "; "

            # NOTE: don't use "= if (...) { ... } else { @() }" here - an empty array
            # returned from an if-block unrolls to $null. Assign explicitly instead.
            $groupTypesArr = @()
            if ($entraGroup.GroupTypes) { $groupTypesArr = @($entraGroup.GroupTypes) }
            $members = Get-GroupMembersWithDetails -GroupId $entraGroup.Id -GroupDisplayName $entraGroup.DisplayName -UserDetailsLookup $userDetailsLookup
            $allGroupMembers += $members
            $memberCount = @($members).Count

            $welcomeEnabled = Get-GroupWelcomeEmailEnabled -GroupId $entraGroup.Id -GroupTypes $groupTypesArr
            $welcomeStatus  = if ($null -eq $welcomeEnabled) { 'N/A' }
                              elseif ($welcomeEnabled)       { 'Yes' }
                              else                           { 'No' }

            $groupWelcomeMail += [pscustomobject]@{
                GroupId             = $entraGroup.Id
                GroupDisplayName    = $entraGroup.DisplayName
                GroupType           = $groupCategory
                WelcomeEmailEnabled = $welcomeStatus
                MemberCount         = $memberCount
            }

            $allGroupsCsv += [pscustomobject]@{
                Id                                   = $entraGroup.Id
                DisplayName                          = $entraGroup.DisplayName
                Description                          = $entraGroup.Description
                Source                               = $source
                GroupCategory                        = $groupCategory
                CreatedDateTime                      = $entraGroup.CreatedDateTime
                GroupTypes                           = ($entraGroup.GroupTypes -join "; ")
                MailEnabled                          = $entraGroup.MailEnabled
                SecurityEnabled                      = $entraGroup.SecurityEnabled
                MailNickname                         = $entraGroup.MailNickname
                Visibility                           = $entraGroup.Visibility
                IsAssignableToRole                   = $entraGroup.IsAssignableToRole
                DynamicGroupMembershipRule           = $entraGroup.MembershipRule
                DynamicMembershipRuleProcessingState = $entraGroup.MembershipRuleProcessingState
                ResourceProvisioningOptions          = ($entraGroup.ResourceProvisioningOptions -join "; ")
                OnPremisesSyncEnabled                = $entraGroup.OnPremisesSyncEnabled
                Owners                               = $ownerUpns
                Licenses                             = $licenses
                RolesAndAdministrators               = $roleAssignments
                AdministrativeUnits                  = $memberOfInfo.AdministrativeUnits
                GroupMemberships                     = $memberOfInfo.GroupMemberships
                Applications                         = $applications
                MemberCount                          = $memberCount
                WelcomeEmailEnabled                  = $welcomeStatus
            }
        }

        Write-Host "Export group summary information to files..."
        Export-InventoryData -Data $groupBasicInfo -BaseName 'EntraGroupsBasicInfo' -WorksheetName 'GroupsBasicInfo' -Paths $Paths

        Write-Host "Export group data to files..."
        Export-InventoryData -Data $allGroupsCsv -BaseName 'EntraGroups' -Paths $Paths

        Write-Host "Export group members ($(@($allGroupMembers).Count) rows) to files..."
        Export-InventoryData -Data $allGroupMembers -BaseName 'GroupMembers' -Paths $Paths

        Write-Host "Export group welcome-email status ($(@($groupWelcomeMail).Count) groups) to files..."
        Export-InventoryData -Data $groupWelcomeMail -BaseName 'GroupWelcomeEmail' -Paths $Paths

        Write-Host "Done. Groups Information collected: $(@($allGroupsCsv).Count)" -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green

        return @{
            GroupBasicInfo    = $groupBasicInfo
            EntraGroups       = $allGroupsCsv
            GroupMembers      = $allGroupMembers
            GroupWelcomeEmail = $groupWelcomeMail
        }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}

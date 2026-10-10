@{
    # SKU part number  ->  text that appears in the icon file name
    #
    # The map builder looks for an icon in three passes, stopping at the first hit:
    #   1. a file named exactly after the SKU        ENTERPRISEPACK.png
    #   2. this table, matched anywhere in the name  10231-icon-service-Entra-ID-Protection.png
    #   3. the SKU up to its first underscore        SPE_E3 -> SPE
    #
    # Matching is case-insensitive and matches part of the name, so a short
    # distinctive fragment is enough. After adding icons, run Show-IconMapStatus
    # to see what still needs an entry.

    Map = @{
        # --- Microsoft 365 / Office suites -> Entra ID / Licenses ---
        'ENTERPRISEPACK'                  = 'Entra-Identity-Licenses'
        'SPE_E3'                          = 'Entra-Identity-Licenses'
        'SPE_E5'                          = 'Entra-Identity-Licenses'
        'SPE_F1'                          = 'Entra-Identity-Licenses'
        'Microsoft_365_E3_Extra_Features' = 'Entra-Identity-Licenses'
        'STANDARDWOFFPACK_FACULTY'        = 'Entra-Identity-Licenses'
        'STANDARDWOFFPACK_STUDENT'        = 'Entra-Identity-Licenses'
        'M365EDU_A3_FACULTY'              = 'Entra-Identity-Licenses'
        'M365EDU_A3_STUUSEBNFT'           = 'Entra-Identity-Licenses'

        # --- identity and security ---
        'EMS'                             = 'Enterprise-Applications'
        'AAD_PREMIUM'                     = 'Entra-ID-Protection'
        'AAD_PREMIUM_P2'                  = 'Entra-ID-Protection'
        'IDENTITY_THREAT_PROTECTION'      = 'Identity-Secure-Score'
        'DEFENDER_ENDPOINT_P1'            = 'Microsoft-Defender-for-Cloud'
        'RMSBASIC'                        = 'Information-Protection'

        # --- workloads ---
        'EXCHANGESTANDARD'                = 'Global-Secure-Access'
        'WIN10_VDA_E5'                    = 'Multi-Factor-Authentication'
        'Teams_Premium_(for_Departments)' = 'Enterprise-Applications'
        'STREAM'                          = 'Enterprise-Applications'

        # --- Power Platform (no dedicated icons in this set -> Application group) ---
        'POWER_BI_PRO'                    = 'Application-Security-Groups'
        'POWER_BI_STANDARD'               = 'Application-Security-Groups'
        'FLOW_FREE'                       = 'API-Proxy'
        'POWERAPPS_VIRAL'                 = 'application-group'
        'POWERAPPS_DEV'                   = 'application-group'
        'POWERAPPS_PER_USER'             = 'application-group'
        'Power_Pages_vTrial_for_Makers'   = 'application-group'
        'CCIBOTS_PRIVPREV_VIRAL'          = 'Enterprise-Applications'
        'SPZA_IW'                         = 'app-registrations'

        # --- desktop / device ---
        'VISIOCLIENT'                     = 'app-registrations'
        'VISIOONLINE_PLAN1'               = 'app-registrations'
        'PROJECTPREMIUM'                  = 'app-registrations'
        'OFFICE_PROPLUS_DEVICE1'          = 'device-compliance'

        # --- AI ---
        'Microsoft_365_Copilot'           = 'Code'

        # --- map block types (groups, apps, CA, owners) -> Entra pack fragments ---
        'Azure-Active-Directory'          = 'Microsoft-Entra-ID'
        'Tenant'                          = 'Microsoft-Entra-ID'
        'Group'                           = 'Groups'
        'SecurityGroup'                   = 'Groups'
        'M365Group'                       = 'Groups'
        'Owner'                           = 'User'
        'User'                            = 'User'
        'ConditionalAccess'               = 'Conditional-Access'
        'AppRegistration'                 = 'App-Registrations'
        'EnterpriseApp'                   = 'Enterprise-Applications'
        'Workplace'                       = 'Microsoft-Entra-ID'
          # --- hub/overview service names -> fragments matching the bundled icons ---
          '10230-icon-service-Users'                                   = 'users'
          '10233-icon-service-Conditional-Access'                      = 'conditional-access'
          '10223-icon-service-Groups'                                  = 'groups'
          '10223-icon-service-GroupOwner'                              = 'person'
          '10225-icon-service-Enterprise-Applications'                 = 'enterprise-applications'
          '10340-icon-service-Entra-Identity-Roles-and-Administrators' = 'shield'
          '10222-icon-service-Entra-Domain-Services'                   = 'building'
          'sspr-user-lock'                                             = 'lock'
    }
}

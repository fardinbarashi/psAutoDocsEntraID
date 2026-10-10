@{
    # Groups purchased SKUs for the licence map so related licences sit together.
    #
    # Order matters: groups are laid out left to right in the order listed here.
    # Put platform/app groups first (they land on the left) and suites last
    # (they land on the right). Each SKU is matched by an exact part number or,
    # if none matches, by the first fragment found in its name.
    #
    # Anything not matched falls into 'Other', which is placed last.

    Groups = @(
        @{
            Name = 'Platform & apps'
            Skus = @('POWER_BI_PRO','POWER_BI_STANDARD','POWERAPPS_PER_USER','POWERAPPS_VIRAL',
                     'POWERAPPS_DEV','FLOW_FREE','VISIOCLIENT','VISIOONLINE_PLAN1','PROJECTPREMIUM')
            Match = @('POWER_BI','POWERAPPS','FLOW','VISIO','PROJECT')
        }
        @{
            Name = 'Security & identity'
            Skus = @('EMS','DEFENDER_ENDPOINT_P1','IDENTITY_THREAT_PROTECTION','RMSBASIC',
                     'AAD_PREMIUM','AAD_PREMIUM_P2')
            Match = @('EMS','DEFENDER','IDENTITY','RMS','AAD_PREMIUM')
        }
        @{
            Name = 'Workloads'
            Skus = @('EXCHANGESTANDARD','WIN10_VDA_E5','Teams_Premium_(for_Departments)',
                     'STREAM','Microsoft_365_Copilot')
            Match = @('EXCHANGE','WIN10','TEAMS','STREAM','COPILOT')
        }
        @{
            Name = 'Suites & education'
            Skus = @('ENTERPRISEPACK','SPE_E3','SPE_E5','SPE_F1','Microsoft_365_E3_Extra_Features',
                     'M365EDU_A3_FACULTY','M365EDU_A3_STUUSEBNFT',
                     'STANDARDWOFFPACK_FACULTY','STANDARDWOFFPACK_STUDENT')
            Match = @('ENTERPRISEPACK','SPE_','M365EDU','STANDARDWOFFPACK','Microsoft_365_E3')
        }
    )
}

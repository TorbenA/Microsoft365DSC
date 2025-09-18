<#
This example is used to test new resources and showcase the usage of new resources being worked on.
It is not meant to use as a production baseline.
#>

Configuration Example
{
    param(
        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint
    )
    Import-DscResource -ModuleName Microsoft365DSC
    node localhost
    {
        AADCrossTenantIdentitySyncPolicyPartner "AADCrossTenantIdentitySyncPolicyPartner-Fabrikam"
        {
            ApplicationId                                       = $ApplicationId;
            CertificateThumbprint                               = $CertificateThumbprint;
            CrossTenantAccessPolicyConfigurationPartnerTenantId = "d8295cae-8bd0-4a7f-9288-933d2dc4573c";
            DisplayName                                         = "IdentitySync";
            Ensure                                              = "Absent";
            IsSyncAllowed                                       = $True;
            TenantId                                            = $TenantId;
        }
    }
}

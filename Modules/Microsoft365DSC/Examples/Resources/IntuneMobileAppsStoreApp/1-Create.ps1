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
        IntuneMobileAppsStoreApp "IntuneMobileAppsStoreApp-Store App"
        {
            TargetPlatform        = "iOS"
            ApplicableDeviceType  = MSFT_MicrosoftGraphiosDeviceType{
                iPad = $True
                iPhoneAndIPod = $True
            }
            AppStoreUrl           = "https://itunes.apple.com/us/app/store-app/id1087422156?mt=8"
            BundleId              = "com.contoso.storeapp"
            Description           = "Store App Description";
            Developer             = "Contoso";
            DisplayName           = "Store App";
            Ensure                = "Present";
            InformationUrl        = "";
            IsFeatured            = $False;
            Notes                 = "";
            Owner                 = "";
            PrivacyInformationUrl = "";
            Publisher             = "Contoso";
            Assignments          = @(
                MSFT_DeviceManagementMobileAppAssignment {
                    groupDisplayName = 'All devices'
                    deviceAndAppManagementAssignmentFilterType = 'none'
                    dataType = '#microsoft.graph.allDevicesAssignmentTarget'
                    intent = 'required'
                }
            );
            Categories             = @(
                MSFT_DeviceManagementMobileAppCategory{
                    Id = "2185c6bf-1b3d-4daa-a0bc-79cb4fad9c87"
                    DisplayName = "App Category 1"
                }
            );
            ApplicationId         = $ApplicationId;
            TenantId              = $TenantId;
            CertificateThumbprint = $CertificateThumbprint;
        }
    }
}

<#
This example creates a Terms of Use Agreement that requires re-acceptance every 30 days on each device.
#>

Configuration Example
{
    param(
        [Parameter(Mandatory = $true)]
        [PSCredential]
        $Credential
    )
    Import-DscResource -ModuleName Microsoft365DSC

    node localhost
    {
        AADAgreement 'MonthlyDeviceTermsOfUse'
        {
            DisplayName                          = "Monthly Device Terms of Use"
            IsViewingBeforeAcceptanceRequired    = $true
            IsPerDeviceAcceptanceRequired        = $true
            UserReacceptRequiredFrequency        = "P30D"
            AcceptanceStatement                  = "I have read and accept the terms of use for this device"
            FileData                             = "TERMS OF USE FOR DEVICE ACCESS\n\nBy accepting these terms, you agree to comply with all company policies..."
            FileName                             = "device_terms.txt"
            Language                             = "en-US"
            Ensure                               = "Present"
            Credential                           = $Credential
        }
    }
}
<#
This example removes an existing Azure AD Terms of Use Agreement.
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
        AADAgreement 'CompanyTermsOfUse'
        {
            DisplayName = "Company Terms of Use"
            Ensure      = "Absent"
            Credential  = $Credential
        }
    }
}
[hashtable]$Script:M365DSCTelemetryConnectionToGraphParams = @{}

<#
.DESCRIPTION
    This function gets all resources that support the specified authentication method and
    determines the most secure authentication method supported by the resource.

.FUNCTIONALITY
    Internal
#>
function Get-M365DSCComponentsWithMostSecureAuthenticationType
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param
    (
        [Parameter()]
        [System.String[]]
        [ValidateSet('ApplicationWithSecret', 'CertificateThumbprint', 'CertificatePath', 'Credentials', 'CredentialsWithTenantId', 'CredentialsWithApplicationId', 'ManagedIdentity', 'AccessTokens')]
        $AuthenticationMethod,

        [Parameter()]
        [System.String[]]
        $Resources
    )

    $modules = Get-ChildItem -Path ($PSScriptRoot + '\..\DSCResources\') -Recurse -Filter '*.psm1'
    $Components = @()
    foreach ($resource in $modules)
    {
        if ($Resources -contains ($resource.Name -replace '.psm1', '' -replace 'MSFT_', ''))
        {
            Import-Module $resource.FullName -Force -Function @('Set-TargetResource') -DisableNameChecking
            $parameters = (Get-Command 'Set-TargetResource').Parameters.Keys

            # Case - Resource supports CertificateThumbprint
            if ($AuthenticationMethod.Contains('CertificateThumbprint') -and `
                    $parameters.Contains('ApplicationId') -and `
                    $parameters.Contains('CertificateThumbprint') -and `
                    $parameters.Contains('TenantId'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'CertificateThumbprint'
                }
            }

            # Case - Resource supports CertificatePath
            elseif ($AuthenticationMethod.Contains('CertificatePath') -and `
                    $parameters.Contains('ApplicationId') -and `
                    $parameters.Contains('CertificatePath') -and `
                    $parameters.Contains('TenantId'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'CertificatePath'
                }
            }

            # Case - Resource supports ApplicationSecret
            elseif ($AuthenticationMethod.Contains('ApplicationWithSecret') -and `
                    $parameters.Contains('ApplicationId') -and `
                    $parameters.Contains('ApplicationSecret') -and `
                    $parameters.Contains('TenantId'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'ApplicationSecret'
                }
            }
            # Case - Resource supports CredentialWithTenantId
            elseif ($AuthenticationMethod.Contains('CredentialsWithTenantId') -and `
                    $parameters.Contains('Credential') -and $parameters.Contains('TenantId') -and `
                    -not $resource.Name.StartsWith('MSFT_SPO') -and `
                    -not $resource.Name.StartsWith('MSFT_OD') -and `
                    -not $resource.Name.StartsWith('MSFT_PP'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'CredentialsWithTenantId'
                }
            }
            # Case - Resource supports Credential using CredentialsWithApplicationId
            elseif ($AuthenticationMethod.Contains('CredentialsWithApplicationId') -and `
                    $parameters.Contains('Credential'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'CredentialsWithApplicationId'
                }
            }
            # Case - Resource supports Credential
            elseif ($AuthenticationMethod.Contains('Credentials') -and `
                    $parameters.Contains('Credential'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'Credentials'
                }
            }
            elseif ($AuthenticationMethod.Contains('ManagedIdentity') -and `
                    $parameters.Contains('ManagedIdentity'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'ManagedIdentity'
                }
            }
            elseif ($AuthenticationMethod.Contains('AccessTokens') -and `
                    $parameters.Contains('AccessTokens'))
            {
                $Components += @{
                    Resource   = $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
                    AuthMethod = 'AccessTokens'
                }
            }
        }
    }
    return $Components
}

<#
.DESCRIPTION
    This function creates a new connection to the specified M365 workload

.FUNCTIONALITY
    Internal
#>
function New-M365DSCConnection
{
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'DefenderForEndpoint', 'EngageHub', 'ExchangeOnline', 'Fabric', 'Licensing', `
                'SecurityComplianceCenter', 'PnP', 'PowerPlatforms', 'PowerPlatformREST', `
                'MicrosoftTeams', 'MicrosoftGraph', 'SharePointOnlineREST', 'Tasks')]
        [System.String]
        $Workload,

        [Parameter(Mandatory = $true)]
        [ValidateScript({
                if ($null -ne $_.Credential)
                {
                    $invalid = $_.Credential.Username -notmatch '.onmicrosoft.'
                    if (-not $invalid)
                    {
                        return $true
                    }
                    else
                    {
                        Write-Warning -Message 'We recommend providing the username in the format of <tenant>.onmicrosoft.* for the Credential property.'
                    }
                }

                if ($null -ne $_.TenantId)
                {
                    $parseGuid = [System.Guid]::Empty
                    $isValid = [System.Guid]::TryParse($_.TenantId, [ref]$parseGuid)
                    if ($isValid)
                    {
                        throw 'Please provide the tenant name (e.g., contoso.onmicrosoft.com) for TenantId instead of its GUID.'
                    }

                    $isValid = $_.TenantId -match '.onmicrosoft.'
                    if ($isValid)
                    {
                        return $true
                    }
                    else
                    {
                        Write-Warning -Message 'We recommend providing the tenant name in format <tenant>.onmicrosoft.* for TenantId.'
                    }
                }
                return $true
            })]
        [System.Collections.Hashtable]
        $InboundParameters,

        [Parameter()]
        [System.String]
        $Url,

        [Parameter()]
        [switch]
        $EnableSearchOnlySession
    )

    if (-not (Test-IsM365DSCRequiredModulesLoaded))
    {
        $requiredModules = Get-M365DSCRequiredModules
        foreach ($requiredModule in $requiredModules)
        {
            Write-Verbose -Message "Ensuring required module '$requiredModule' is loaded."
            Confirm-M365DSCLoadedModule -ModuleName $requiredModule
        }
        Set-M365DSCRequiredModulesLoaded -Value $true
    }

    if ($Workload -eq 'MicrosoftTeams')
    {
        try
        {
            $null = Get-Command 'Connect-MicrosoftTeams' -ErrorAction Stop
        }
        catch
        {
            Import-Module 'MicrosoftTeams' -Global -Force -Alias @() -Cmdlet @() -Variable @() -DisableNameChecking | Out-Null
        }
    }

    Write-Verbose -Message "Attempting connection to {$Workload} with:"
    Write-Verbose -Message "$($InboundParameters | Out-String)"

    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add('Source', 'M365DSCUtil')
    $data.Add('Workload', $Workload)

    $Script:M365DSCTelemetryConnectionToGraphParams = @{}

    # Keep track of workloads we already connected so that we don't send additional Telemetry events.
    if ($null -eq $Script:M365ConnectedToWorkloads)
    {
        Write-Verbose -Message 'Initializing the Connected To Workloads List.'
        $Script:M365ConnectedToWorkloads = @()
    }

    # Convert ApplicationSecret from SecureString to plain string.
    if ($InboundParameters.ApplicationSecret)
    {
        $InboundParameters.ApplicationSecret = $InboundParameters.ApplicationSecret.GetNetworkCredential().Password
    }

    #region Validation
    if ($null -ne $InboundParameters.Credential -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint))
    {
        $message = 'Both Authentication methods are attempted'
        Write-Verbose -Message $message
        $data.Add('Exception', $message)
        $errorText = "You can't specify both the Credential and CertificateThumbprint"
        $data.Add('CustomMessage', $errorText)
        Add-M365DSCTelemetryEvent -Type 'Error' -Data $data
        throw $errorText
    }

    if ($null -eq $InboundParameters.Credential -and `
            [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint) -and `
            -not $InboundParameters.ManagedIdentity -and `
            $null -eq $InboundParameters.AccessTokens)
    {
        $message = 'No Authentication method was provided'
        Write-Verbose -Message $message
        $message += "`r`nProvided Keys --> $($InboundParameters.Keys)"
        $data.Add('Exception', $message)
        $errorText = 'You must specify either the Credential or ApplicationId, TenantId and CertificateThumbprint parameters.'
        $data.Add('CustomMessage', $errorText)
        Add-M365DSCTelemetryEvent -Type 'Error' -Data $data
        throw $errorText
    }
    #endregion Validation

    # Determine connection mode using the shared helper.
    $connectionMode = Get-M365DSCAuthenticationMode -Parameters $InboundParameters
    if ($connectionMode -eq 'Interactive')
    {
        throw 'Could not determine authentication method'
    }
    Write-Verbose -Message "Connecting via $connectionMode"

    #region Build Connect-M365Tenant splat
    $connectParams = @{
        Workload                = $Workload
        EnableSearchOnlySession = $EnableSearchOnlySession.IsPresent
    }

    if (-not [System.String]::IsNullOrEmpty($Url))
    {
        $connectParams.Url = $Url
    }

    switch ($connectionMode)
    {
        'Credentials'
        {
            $connectParams.Credential = $InboundParameters.Credential
        }
        'CredentialsWithApplicationId'
        {
            $connectParams.ApplicationId = $InboundParameters.ApplicationId
            $connectParams.Credential = $InboundParameters.Credential
        }
        'CredentialsWithTenantId'
        {
            $connectParams.TenantId = $InboundParameters.TenantId
            $connectParams.Credential = $InboundParameters.Credential
        }
        'ServicePrincipalWithPath'
        {
            $connectParams.ApplicationId = $InboundParameters.ApplicationId
            $connectParams.TenantId = $InboundParameters.TenantId
            $connectParams.CertificatePassword = $InboundParameters.CertificatePassword.Password
            $connectParams.CertificatePath = $InboundParameters.CertificatePath
        }
        'ServicePrincipalWithSecret'
        {
            $connectParams.ApplicationId = $InboundParameters.ApplicationId
            $connectParams.TenantId = $InboundParameters.TenantId
            $connectParams.ApplicationSecret = $InboundParameters.ApplicationSecret
        }
        'ServicePrincipalWithThumbprint'
        {
            $connectParams.ApplicationId = $InboundParameters.ApplicationId
            $connectParams.TenantId = $InboundParameters.TenantId
            $connectParams.CertificateThumbprint = $InboundParameters.CertificateThumbprint
        }
        'ManagedIdentity'
        {
            $connectParams.Identity = $true
            $connectParams.TenantId = $InboundParameters.TenantId
        }
        'AccessTokens'
        {
            $connectParams.AccessTokens = $InboundParameters.AccessTokens
            $connectParams.TenantId = $InboundParameters.TenantId
        }
    }
    #endregion

    Connect-M365Tenant @connectParams

    #region Update telemetry cache
    $telemetryCacheKeys = switch ($connectionMode)
    {
        'Credentials'                    { @('Credential') }
        'CredentialsWithApplicationId'   { @('Credential', 'ApplicationId') }
        'CredentialsWithTenantId'        { @('Credential', 'TenantId') }
        'ServicePrincipalWithPath'       { @('ApplicationId', 'TenantId', 'CertificatePath') }
        'ServicePrincipalWithSecret'     { @('ApplicationId', 'TenantId', 'ApplicationSecret') }
        'ServicePrincipalWithThumbprint' { @('ApplicationId', 'TenantId', 'CertificateThumbprint') }
        'ManagedIdentity'                { @('TenantId') }
        'AccessTokens'                   { @('AccessTokens', 'TenantId') }
    }

    foreach ($key in $telemetryCacheKeys)
    {
        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey($key) -and
            $null -ne $InboundParameters[$key])
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add($key, $InboundParameters[$key])
        }
    }

    # Handle special telemetry cache values not directly from InboundParameters.
    if ($connectionMode -eq 'ManagedIdentity' -and
        -not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('Identity'))
    {
        $Script:M365DSCTelemetryConnectionToGraphParams.Add('Identity', $true)
    }
    if ($connectionMode -eq 'ServicePrincipalWithPath' -and
        -not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('CertificatePassword'))
    {
        $Script:M365DSCTelemetryConnectionToGraphParams.Add('CertificatePassword', $InboundParameters.CertificatePassword.Password)
    }
    #endregion

    #region Emit connection telemetry
    # The Credentials mode uses 'Credential' (no trailing 's') as tracking key for backward compatibility.
    $trackingKey = if ($connectionMode -eq 'Credentials') { 'Credential' } else { $connectionMode }
    $workloadTrackingKey = "$Workload-$trackingKey"

    if (-not ($Script:M365ConnectedToWorkloads -contains $workloadTrackingKey))
    {
        $data.Add('ConnectionMode', $connectionMode)

        if (-not $data.ContainsKey('Tenant'))
        {
            if (-not [System.String]::IsNullOrEmpty($InboundParameters.TenantId))
            {
                $data.Add('Tenant', $InboundParameters.TenantId)
            }
            elseif ($null -ne $InboundParameters.Credential)
            {
                try
                {
                    $tenantId = $InboundParameters.Credential.Username.Split('@')[1]
                    $data.Add('Tenant', $tenantId)
                    if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
                    {
                        $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $tenantId)
                    }
                }
                catch
                {
                    Write-Verbose -Message $_
                }
            }
        }

        Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
        $Script:M365ConnectedToWorkloads += $workloadTrackingKey
    }
    #endregion

    return $connectionMode
}

<#
.DESCRIPTION
    This function gets the used authentication mode based on the specified parameters

.FUNCTIONALITY
    Internal
#>
function Get-M365DSCAuthenticationMode
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Parameters
    )

    # Cache frequently accessed values to reduce hashtable lookups
    $applicationId = $Parameters.ApplicationId
    $tenantId = $Parameters.TenantId
    $credential = $Parameters.Credential

    # Check service principal authentication modes first (most common in automation)
    if ($applicationId -and $tenantId)
    {
        if ($Parameters.CertificateThumbprint)
        {
            return 'ServicePrincipalWithThumbprint'
        }
        if ($Parameters.ApplicationSecret)
        {
            return 'ServicePrincipalWithSecret'
        }
        if ($Parameters.CertificatePath -and $Parameters.CertificatePassword)
        {
            return 'ServicePrincipalWithPath'
        }
    }

    # Check credential-based authentication
    if ($credential)
    {
        if ($applicationId)
        {
            return 'CredentialsWithApplicationId'
        }
        if ($tenantId)
        {
            return 'CredentialsWithTenantId'
        }
        return 'Credentials'
    }

    # Check other authentication modes
    if ($Parameters.ManagedIdentity)
    {
        return 'ManagedIdentity'
    }

    if ($Parameters.AccessTokens)
    {
        return 'AccessTokens'
    }

    # Default to interactive
    return 'Interactive'
}

<#
.DESCRIPTION
    This function retrieves the telemetry connection parameters for the current session.

.FUNCTIONALITY
    Internal.
#>
function Get-M365DSCTelemetryConnectionParameter
{
    [CmdletBinding()]
    param ()

    $Script:M365DSCTelemetryConnectionToGraphParams.Clone()
}

<#
.DESCRIPTION
    This function sets the telemetry connection parameters for the current session.

.FUNCTIONALITY
    Internal.
#>
function Set-M365DSCTelemetryConnectionParameter
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    $Script:M365DSCTelemetryConnectionToGraphParams = $Parameters.Clone()
}

<#
.DESCRIPTION
    This function returns the connection workloads for the specified DSC resources

.PARAMETER ResourceNames
    Specifies the resources for which the connection workloads should be determined.
    Either a single string, an array of strings or an object with 'Name' and 'AuthenticationMethod' can be provided.

.EXAMPLE
    PS> Get-M365DSCConnectedWorkloadList -ResourceNames AADUser

.EXAMPLE
    PS> Get-M365DSCConnectedWorkloadList -ResourceNames @('AADUser', 'AADGroup')

.EXAMPLE
    PS> Get-M365DSCConnectedWorkloadList -ResourceNames @{Name = 'AADUser'; AuthenticationMethod = 'Credentials'}

.FUNCTIONALITY
    Public
#>
function Get-M365DSCConnectedWorkloadList
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable[]])]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [System.Array]
        $ResourceNames
    )

    [Array] $workloads = @()
    foreach ($resource in $ResourceNames)
    {
        $resourceName = $resource.Name
        $authMethod = $resource.AuthenticationMethod
        if ([System.String]::IsNullOrEmpty($resourceName))
        {
            $resourceName = $resource
        }
        switch ($resourceName.Substring(0, 2).ToUpper())
        {
            'AA'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('MicrosoftGraph'))
                {
                    $workloads += @{
                        Name                 = 'MicrosoftGraph'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'EX'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('ExchangeOnline'))
                {
                    $workloads += @{
                        Name                 = 'ExchangeOnline'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'In'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('MicrosoftGraph'))
                {
                    $workloads += @{
                        Name                 = 'MicrosoftGraph'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'O3'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('MicrosoftGraph') -and $resource.Name -eq 'O365Group')
                {
                    $workloads += @{
                        Name                 = 'MicrosoftGraph'
                        AuthenticationMethod = $authMethod
                    }
                }
                elseif (-not $workloads.Name -or -not $workloads.Name.Contains('ExchangeOnline'))
                {
                    $workloads += @{
                        Name                 = 'ExchangeOnline'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'OD'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('PnP'))
                {
                    $workloads += @{
                        Name                 = 'PnP'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'Pl'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('MicrosoftGraph'))
                {
                    $workloads += @{
                        Name                 = 'MicrosoftGraph'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'SP'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('PnP'))
                {
                    $workloads += @{
                        Name                 = 'PnP'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'SC'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('SecurityComplianceCenter'))
                {
                    $workloads += @{
                        Name                 = 'SecurityComplianceCenter'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
            'Te'
            {
                if (-not $workloads.Name -or -not $workloads.Name.Contains('MicrosoftTeams'))
                {
                    $workloads += @{
                        Name                 = 'MicrosoftTeams'
                        AuthenticationMethod = $authMethod
                    }
                }
            }
        }
    }
    return ($workloads | Sort-Object { $_.Name })
}

Export-ModuleMember -Function @(
    'Get-M365DSCAuthenticationMode',
    'Get-M365DSCComponentsWithMostSecureAuthenticationType',
    'Get-M365DSCConnectedWorkloadList',
    'Get-M365DSCTelemetryConnectionParameter',
    'New-M365DSCConnection',
    'Set-M365DSCTelemetryConnectionParameter'
)

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        #region resource generator code
        [Parameter()]
        [System.String]
        $AdDomainName,

        [Parameter()]
        [System.String]
        $AdDomainPassword,

        [Parameter()]
        [System.String]
        $AdDomainUsername,

        [Parameter()]
        [System.String]
        $AlternateResourceUrl,

        [Parameter()]
        [ValidateSet('hybridAzureADJoin','azureADJoin')]
        [System.String]
        $ConnectionType,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.Boolean]
        $HealthCheckPaused,

        [Parameter()]
        [ValidateSet('pending','running','passed','failed','warning','informational')]
        [System.String]
        $HealthCheckStatus,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $HealthCheckStatusDetail,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $HealthCheckStatusDetails,

        [Parameter()]
        [System.Boolean]
        $InUse,

        [Parameter()]
        [System.Boolean]
        $InUseByCloudPc,

        [Parameter()]
        [ValidateSet('windows365','devBox','rpaBox')]
        [System.String]
        $ManagedBy,

        [Parameter()]
        [System.String]
        $OrganizationalUnit,

        [Parameter()]
        [System.String]
        $ResourceGroupId,

        [Parameter()]
        [System.String[]]
        $ScopeIds,

        [Parameter()]
        [System.String]
        $SubnetId,

        [Parameter()]
        [System.String]
        $SubscriptionId,

        [Parameter()]
        [System.String]
        $SubscriptionName,

        [Parameter()]
        [ValidateSet('hybridAzureADJoin','azureADJoin')]
        [System.String]
        $Type,

        [Parameter()]
        [System.String]
        $VirtualNetworkId,

        [Parameter()]
        [System.String]
        $VirtualNetworkLocation,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Id,

        #endregion

        [Parameter()]
        [System.String]
        [ValidateSet('Absent', 'Present')]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    Write-Verbose -Message "Getting configuration for the Intune Azure Network Connection for Windows365 with Id {$Id} and DisplayName {$DisplayName}"

    try
    {
        if (-not $Script:exportedInstance -or $Script:exportedInstance.DisplayName -ne $DisplayName)
        {
            $null = New-M365DSCConnection -Workload 'MicrosoftGraph' `
                -InboundParameters $PSBoundParameters

            #Ensure the proper dependencies are installed in the current environment.
            Confirm-M365DSCDependencies

            #region Telemetry
            $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
            $CommandName = $MyInvocation.MyCommand
            $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
                -CommandName $CommandName `
                -Parameters $PSBoundParameters
            Add-M365DSCTelemetryEvent -Data $data
            #endregion

            $nullResult = $PSBoundParameters
            $nullResult.Ensure = 'Absent'

            $getValue = $null

            #region resource generator code
            if (-not [System.String]::IsNullOrEmpty($Id))
            {
                $getValue = Get-MgBetaDeviceManagementVirtualEndpointOnPremiseConnection -CloudPcOnPremisesConnectionId $Id  -ErrorAction SilentlyContinue
            }

            if ($null -eq $getValue)
            {
                Write-Verbose -Message "Could not find an Intune Azure Network Connection for Windows365 with Id {$Id}"

                if (-not [System.String]::IsNullOrEmpty($DisplayName))
                {
                    $getValue = Get-MgBetaDeviceManagementVirtualEndpointOnPremiseConnection `
                        -Filter "DisplayName eq '$($DisplayName -replace "'", "''")'" `
                        -ErrorAction SilentlyContinue | Where-Object `
                        -FilterScript {
                            $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.CloudPcOnPremisesConnection"
                        }
                }
            }
            #endregion
            if ($null -eq $getValue)
            {
                Write-Verbose -Message "Could not find an Intune Azure Network Connection for Windows365 with DisplayName {$DisplayName}."
                return $nullResult
            }
        }
        else
        {
            $getValue = $Script:exportedInstance
        }
        $Id = $getValue.Id
        Write-Verbose -Message "An Intune Azure Network Connection for Windows365 with Id {$Id} and DisplayName {$DisplayName} was found"

        #region resource generator code
        $complexHealthCheckStatusDetail = @{}
        if ($null -ne $getValue.HealthCheckStatusDetail.endDateTime)
        {
            $complexHealthCheckStatusDetail.Add('EndDateTime', ([DateTimeOffset]$getValue.HealthCheckStatusDetail.endDateTime).ToString('o'))
        }
        $complexHealthChecks = @()
        foreach ($currentHealthChecks in $getValue.HealthCheckStatusDetail.healthChecks)
        {
            $myHealthChecks = @{}
            $myHealthChecks.Add('AdditionalDetail', $currentHealthChecks.additionalDetail)
            $myHealthChecks.Add('AdditionalDetails', $currentHealthChecks.additionalDetails)
            $myHealthChecks.Add('CorrelationId', $currentHealthChecks.correlationId)
            $myHealthChecks.Add('DisplayName', $currentHealthChecks.displayName)
            if ($null -ne $currentHealthChecks.endDateTime)
            {
                $myHealthChecks.Add('EndDateTime', ([DateTimeOffset]$currentHealthChecks.endDateTime).ToString(''))
            }
            if ($null -ne $currentHealthChecks.errorType)
            {
                $myHealthChecks.Add('ErrorType', $currentHealthChecks.errorType.ToString())
            }
            $myHealthChecks.Add('RecommendedAction', $currentHealthChecks.recommendedAction)
            if ($null -ne $currentHealthChecks.startDateTime)
            {
                $myHealthChecks.Add('StartDateTime', ([DateTimeOffset]$currentHealthChecks.startDateTime).ToString(''))
            }
            if ($null -ne $currentHealthChecks.status)
            {
                $myHealthChecks.Add('Status', $currentHealthChecks.status.ToString())
            }
            if ($myHealthChecks.values.Where({$null -ne $_}).Count -gt 0)
            {
                $complexHealthChecks += $myHealthChecks
            }
        }
        $complexHealthCheckStatusDetail.Add('HealthChecks',$complexHealthChecks)
        if ($null -ne $getValue.HealthCheckStatusDetail.startDateTime)
        {
            $complexHealthCheckStatusDetail.Add('StartDateTime', ([DateTimeOffset]$getValue.HealthCheckStatusDetail.startDateTime).ToString('o'))
        }
        if ($complexHealthCheckStatusDetail.Values.Where({ $null -ne $_ }).Count -eq 0)
        {
            $complexHealthCheckStatusDetail = $null
        }

        $complexHealthCheckStatusDetails = @{}
        if ($null -ne $getValue.HealthCheckStatusDetails.endDateTime)
        {
            $complexHealthCheckStatusDetails.Add('EndDateTime', ([DateTimeOffset]$getValue.HealthCheckStatusDetails.endDateTime).ToString('o'))
        }
        $complexHealthChecks = @()
        foreach ($currentHealthChecks in $getValue.HealthCheckStatusDetails.healthChecks)
        {
            $myHealthChecks = @{}
            $myHealthChecks.Add('AdditionalDetail', $currentHealthChecks.additionalDetail)
            $myHealthChecks.Add('AdditionalDetails', $currentHealthChecks.additionalDetails)
            $myHealthChecks.Add('CorrelationId', $currentHealthChecks.correlationId)
            $myHealthChecks.Add('DisplayName', $currentHealthChecks.displayName)
            if ($null -ne $currentHealthChecks.endDateTime)
            {
                $myHealthChecks.Add('EndDateTime', ([DateTimeOffset]$currentHealthChecks.endDateTime).ToString(''))
            }
            if ($null -ne $currentHealthChecks.errorType)
            {
                $myHealthChecks.Add('ErrorType', $currentHealthChecks.errorType.ToString())
            }
            $myHealthChecks.Add('RecommendedAction', $currentHealthChecks.recommendedAction)
            if ($null -ne $currentHealthChecks.startDateTime)
            {
                $myHealthChecks.Add('StartDateTime', ([DateTimeOffset]$currentHealthChecks.startDateTime).ToString(''))
            }
            if ($null -ne $currentHealthChecks.status)
            {
                $myHealthChecks.Add('Status', $currentHealthChecks.status.ToString())
            }
            if ($myHealthChecks.values.Where({$null -ne $_}).Count -gt 0)
            {
                $complexHealthChecks += $myHealthChecks
            }
        }
        $complexHealthCheckStatusDetails.Add('HealthChecks',$complexHealthChecks)
        if ($null -ne $getValue.HealthCheckStatusDetails.startDateTime)
        {
            $complexHealthCheckStatusDetails.Add('StartDateTime', ([DateTimeOffset]$getValue.HealthCheckStatusDetails.startDateTime).ToString('o'))
        }
        if ($complexHealthCheckStatusDetails.Values.Where({ $null -ne $_ }).Count -eq 0)
        {
            $complexHealthCheckStatusDetails = $null
        }
        #endregion

        #region resource generator code
        $enumConnectionType = $null
        if ($null -ne $getValue.ConnectionType)
        {
            $enumConnectionType = $getValue.ConnectionType.ToString()
        }

        $enumHealthCheckStatus = $null
        if ($null -ne $getValue.HealthCheckStatus)
        {
            $enumHealthCheckStatus = $getValue.HealthCheckStatus.ToString()
        }

        $enumManagedBy = $null
        if ($null -ne $getValue.ManagedBy)
        {
            $enumManagedBy = $getValue.ManagedBy.ToString()
        }

        $enumType = $null
        if ($null -ne $getValue.Type)
        {
            $enumType = $getValue.Type.ToString()
        }
        #endregion

        $results = @{
            #region resource generator code
            AdDomainName             = $getValue.AdDomainName
            AdDomainPassword         = $getValue.AdDomainPassword
            AdDomainUsername         = $getValue.AdDomainUsername
            AlternateResourceUrl     = $getValue.AlternateResourceUrl
            ConnectionType           = $enumConnectionType
            DisplayName              = $getValue.DisplayName
            HealthCheckPaused        = $getValue.HealthCheckPaused
            HealthCheckStatus        = $enumHealthCheckStatus
            HealthCheckStatusDetail  = $complexHealthCheckStatusDetail
            HealthCheckStatusDetails = $complexHealthCheckStatusDetails
            InUse                    = $getValue.InUse
            InUseByCloudPc           = $getValue.InUseByCloudPc
            ManagedBy                = $enumManagedBy
            OrganizationalUnit       = $getValue.OrganizationalUnit
            ResourceGroupId          = $getValue.ResourceGroupId
            ScopeIds                 = $getValue.ScopeIds
            SubnetId                 = $getValue.SubnetId
            SubscriptionId           = $getValue.SubscriptionId
            SubscriptionName         = $getValue.SubscriptionName
            Type                     = $enumType
            VirtualNetworkId         = $getValue.VirtualNetworkId
            VirtualNetworkLocation   = $getValue.VirtualNetworkLocation
            Id                       = $getValue.Id
            Ensure                   = 'Present'
            Credential               = $Credential
            ApplicationId            = $ApplicationId
            TenantId                 = $TenantId
            ApplicationSecret        = $ApplicationSecret
            CertificateThumbprint    = $CertificateThumbprint
            ManagedIdentity          = $ManagedIdentity.IsPresent
            #endregion
        }

        return [System.Collections.Hashtable] $results
    }
    catch
    {
        New-M365DSCLogEntry -Message 'Error retrieving data:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        return $nullResult
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        #region resource generator code
        [Parameter()]
        [System.String]
        $AdDomainName,

        [Parameter()]
        [System.String]
        $AdDomainPassword,

        [Parameter()]
        [System.String]
        $AdDomainUsername,

        [Parameter()]
        [System.String]
        $AlternateResourceUrl,

        [Parameter()]
        [ValidateSet('hybridAzureADJoin','azureADJoin')]
        [System.String]
        $ConnectionType,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.Boolean]
        $HealthCheckPaused,

        [Parameter()]
        [ValidateSet('pending','running','passed','failed','warning','informational')]
        [System.String]
        $HealthCheckStatus,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $HealthCheckStatusDetail,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $HealthCheckStatusDetails,

        [Parameter()]
        [System.Boolean]
        $InUse,

        [Parameter()]
        [System.Boolean]
        $InUseByCloudPc,

        [Parameter()]
        [ValidateSet('windows365','devBox','rpaBox')]
        [System.String]
        $ManagedBy,

        [Parameter()]
        [System.String]
        $OrganizationalUnit,

        [Parameter()]
        [System.String]
        $ResourceGroupId,

        [Parameter()]
        [System.String[]]
        $ScopeIds,

        [Parameter()]
        [System.String]
        $SubnetId,

        [Parameter()]
        [System.String]
        $SubscriptionId,

        [Parameter()]
        [System.String]
        $SubscriptionName,

        [Parameter()]
        [ValidateSet('hybridAzureADJoin','azureADJoin')]
        [System.String]
        $Type,

        [Parameter()]
        [System.String]
        $VirtualNetworkId,

        [Parameter()]
        [System.String]
        $VirtualNetworkLocation,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Id,

        #endregion
        [Parameter()]
        [System.String]
        [ValidateSet('Absent', 'Present')]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    Write-Verbose -Message "Setting configuration of the Intune Azure Network Connection for Windows365 with Id {$Id} and DisplayName {$DisplayName}"

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $currentInstance = Get-TargetResource @PSBoundParameters

    $boundParameters = Remove-M365DSCAuthenticationParameter -BoundParameters $PSBoundParameters


    if ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Absent')
    {
        Write-Verbose -Message "Creating an Intune Azure Network Connection for Windows365 with DisplayName {$DisplayName}"

        $createParameters = ([Hashtable]$boundParameters).Clone()
        $createParameters = Rename-M365DSCCimInstanceParameter -Properties $createParameters
        $createParameters.Remove('Id') | Out-Null

        $keys = (([Hashtable]$createParameters).Clone()).Keys
        foreach ($key in $keys)
        {
            if ($null -ne $createParameters.$key -and $createParameters.$key.GetType().Name -like '*CimInstance*')
            {
                $createParameters.$key = Convert-M365DSCDRGComplexTypeToHashtable -ComplexObject $createParameters.$key
            }
        }
        #region resource generator code
        $createParameters.Add("@odata.type", "#microsoft.graph.CloudPcOnPremisesConnection")
        $policy = New-MgBetaDeviceManagementVirtualEndpointOnPremiseConnection -BodyParameter $createParameters
        #endregion
    }
    elseif ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Updating the Intune Azure Network Connection for Windows365 with Id {$($currentInstance.Id)}"

        $updateParameters = ([Hashtable]$boundParameters).Clone()
        $updateParameters = Rename-M365DSCCimInstanceParameter -Properties $updateParameters

        $updateParameters.Remove('Id') | Out-Null

        $keys = (([Hashtable]$updateParameters).Clone()).Keys
        foreach ($key in $keys)
        {
            if ($null -ne $updateParameters.$key -and $updateParameters.$key.GetType().Name -like '*CimInstance*')
            {
                $updateParameters.$key = Convert-M365DSCDRGComplexTypeToHashtable -ComplexObject $updateParameters.$key
            }
        }

        #region resource generator code
        $updateParameters.Add("@odata.type", "#microsoft.graph.CloudPcOnPremisesConnection")
        Update-MgBetaDeviceManagementVirtualEndpointOnPremiseConnection `
            -CloudPcOnPremisesConnectionId $currentInstance.Id `
            -BodyParameter $UpdateParameters

        #endregion
    }
    elseif ($Ensure -eq 'Absent' -and $currentInstance.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Removing the Intune Azure Network Connection for Windows365 with Id {$($currentInstance.Id)}"
        #region resource generator code
        Remove-MgBetaDeviceManagementVirtualEndpointOnPremiseConnection -CloudPcOnPremisesConnectionId $currentInstance.Id
        #endregion
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        #region resource generator code
        [Parameter()]
        [System.String]
        $AdDomainName,

        [Parameter()]
        [System.String]
        $AdDomainPassword,

        [Parameter()]
        [System.String]
        $AdDomainUsername,

        [Parameter()]
        [System.String]
        $AlternateResourceUrl,

        [Parameter()]
        [ValidateSet('hybridAzureADJoin','azureADJoin')]
        [System.String]
        $ConnectionType,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.Boolean]
        $HealthCheckPaused,

        [Parameter()]
        [ValidateSet('pending','running','passed','failed','warning','informational')]
        [System.String]
        $HealthCheckStatus,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $HealthCheckStatusDetail,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $HealthCheckStatusDetails,

        [Parameter()]
        [System.Boolean]
        $InUse,

        [Parameter()]
        [System.Boolean]
        $InUseByCloudPc,

        [Parameter()]
        [ValidateSet('windows365','devBox','rpaBox')]
        [System.String]
        $ManagedBy,

        [Parameter()]
        [System.String]
        $OrganizationalUnit,

        [Parameter()]
        [System.String]
        $ResourceGroupId,

        [Parameter()]
        [System.String[]]
        $ScopeIds,

        [Parameter()]
        [System.String]
        $SubnetId,

        [Parameter()]
        [System.String]
        $SubscriptionId,

        [Parameter()]
        [System.String]
        $SubscriptionName,

        [Parameter()]
        [ValidateSet('hybridAzureADJoin','azureADJoin')]
        [System.String]
        $Type,

        [Parameter()]
        [System.String]
        $VirtualNetworkId,

        [Parameter()]
        [System.String]
        $VirtualNetworkLocation,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Id,

        #endregion

        [Parameter()]
        [System.String]
        [ValidateSet('Absent', 'Present')]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    Write-Verbose -Message "Testing configuration of the Intune Azure Network Connection for Windows365 with Id {$Id} and DisplayName {$DisplayName}"

    $CurrentValues = Get-TargetResource @PSBoundParameters
    $ValuesToCheck = ([hashtable]$PSBoundParameters).Clone()
    $testResult = $true

    #Compare Cim instances
    foreach ($key in $PSBoundParameters.Keys)
    {
        $source = $PSBoundParameters.$key
        $target = $CurrentValues.$key
        if ($null -ne $source -and $source.GetType().Name -like '*CimInstance*')
        {
            $testResult = Compare-M365DSCComplexObject `
                -Source ($source) `
                -Target ($target)

            if (-not $testResult)
            {
                break
            }

            $ValuesToCheck.Remove($key) | Out-Null
        }
    }

    $ValuesToCheck.Remove('Id') | Out-Null
    $ValuesToCheck = Remove-M365DSCAuthenticationParameter -BoundParameters $ValuesToCheck

    Write-Verbose -Message "Current Values: $(Convert-M365DscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-M365DscHashtableToString -Hashtable $ValuesToCheck)"

    if ($testResult)
    {
        $testResult = Test-M365DSCParameterState -CurrentValues $CurrentValues `
            -Source $($MyInvocation.MyCommand.Source) `
            -DesiredValues $PSBoundParameters `
            -ValuesToCheck $ValuesToCheck.Keys
    }

    Write-Verbose -Message "Test-TargetResource returned $testResult"

    return $testResult
}

function Export-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter()]
        [System.String]
        $Filter,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    $ConnectionMode = New-M365DSCConnection -Workload 'MicrosoftGraph' `
        -InboundParameters $PSBoundParameters

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    try
    {
        #region resource generator code
        [array]$getValue = Get-MgBetaDeviceManagementVirtualEndpointOnPremiseConnection `
            -Filter $Filter `
            -All `
            -ErrorAction Stop
        #endregion

        $i = 1
        $dscContent = ''
        if ($getValue.Length -eq 0)
        {
            Write-M365DSCHost -Message $Global:M365DSCEmojiGreenCheckMark -CommitWrite
        }
        else
        {
            Write-M365DSCHost -Message "`r`n" -DeferWrite
        }
        foreach ($config in $getValue)
        {
            $displayedKey = $config.Id
            if (-not [String]::IsNullOrEmpty($config.displayName))
            {
                $displayedKey = $config.displayName
            }
            elseif (-not [string]::IsNullOrEmpty($config.name))
            {
                $displayedKey = $config.name
            }
            Write-M365DSCHost -Message "    |---[$i/$($getValue.Count)] $displayedKey" -DeferWrite
            $params = @{
                Id                    = $config.Id
                DisplayName           = $config.DisplayName
                Ensure                = 'Present'
                Credential            = $Credential
                ApplicationId         = $ApplicationId
                TenantId              = $TenantId
                ApplicationSecret     = $ApplicationSecret
                CertificateThumbprint = $CertificateThumbprint
                ManagedIdentity       = $ManagedIdentity.IsPresent
                AccessTokens          = $AccessTokens
            }

            $Script:exportedInstance = $config
            $Results = Get-TargetResource @Params
            if ($null -ne $Results.HealthCheckStatusDetail)
            {
                $complexMapping = @(
                    @{
                        Name = 'HealthCheckStatusDetail'
                        CimInstanceName = 'MicrosoftGraphCloudPcOnPremisesConnectionStatusDetail'
                        IsRequired = $False
                    }
                    @{
                        Name = 'HealthChecks'
                        CimInstanceName = 'MicrosoftGraphCloudPcOnPremisesConnectionHealthCheck'
                        IsRequired = $False
                    }
                )
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.HealthCheckStatusDetail `
                    -CIMInstanceName 'MicrosoftGraphcloudPcOnPremisesConnectionStatusDetail' `
                    -ComplexTypeMapping $complexMapping

                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.HealthCheckStatusDetail = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('HealthCheckStatusDetail') | Out-Null
                }
            }
            if ($null -ne $Results.HealthCheckStatusDetails)
            {
                $complexMapping = @(
                    @{
                        Name = 'HealthCheckStatusDetails'
                        CimInstanceName = 'MicrosoftGraphCloudPcOnPremisesConnectionStatusDetails'
                        IsRequired = $False
                    }
                    @{
                        Name = 'HealthChecks'
                        CimInstanceName = 'MicrosoftGraphCloudPcOnPremisesConnectionHealthCheck'
                        IsRequired = $False
                    }
                )
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.HealthCheckStatusDetails `
                    -CIMInstanceName 'MicrosoftGraphcloudPcOnPremisesConnectionStatusDetails' `
                    -ComplexTypeMapping $complexMapping

                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.HealthCheckStatusDetails = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('HealthCheckStatusDetails') | Out-Null
                }
            }

            $currentDSCBlock = Get-M365DSCExportContentForResource -ResourceName $ResourceName `
                -ConnectionMode $ConnectionMode `
                -ModulePath $PSScriptRoot `
                -Results $Results `
                -Credential $Credential `
                -NoEscape @('HealthCheckStatusDetail', 'HealthCheckStatusDetails')
            $dscContent += $currentDSCBlock
            Save-M365DSCPartialExport -Content $currentDSCBlock `
                -FileName $Global:PartialExportFileName
            $i++
            Write-M365DSCHost -Message $Global:M365DSCEmojiGreenCheckMark -CommitWrite
        }
        return $dscContent
    }
    catch
    {
        Write-M365DSCHost -Message $Global:M365DSCEmojiRedX -CommitWrite

        New-M365DSCLogEntry -Message 'Error during Export:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        return ''
    }
}

Export-ModuleMember -Function *-TargetResource

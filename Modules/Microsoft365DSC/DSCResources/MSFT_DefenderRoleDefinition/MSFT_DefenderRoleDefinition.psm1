Confirm-M365DSCModuleDependency -ModuleName 'MSFT_DefenderRoleDefinition'

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Id,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $RolePermissions,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $TenantId,

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

    Write-Verbose -Message "Getting configuration of the Defender Role Definition with Id {$Id} and DisplayName {$DisplayName}"

    try
    {
        if (-not $Script:exportedInstance -or $Script:exportedInstance.DisplayName -ne $DisplayName)
        {
            $null = New-M365DSCConnection -Workload 'MicrosoftGraph' `
                -InboundParameters $PSBoundParameters

            #Ensure the proper dependencies are installed in the current environment.
            Confirm-M365DSCDependencies

            #region Telemetry
            $ResourceName = $MyInvocation.MyCommand.ModuleName -replace 'MSFT_', ''
            $CommandName = $MyInvocation.MyCommand
            $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
                -CommandName $CommandName `
                -Parameters $PSBoundParameters
            Add-M365DSCTelemetryEvent -Data $data
            #endregion

            $nullResult = $PSBoundParameters
            $nullResult.Ensure = 'Absent'

            if (-not [System.String]::IsNullOrEmpty($Id))
            {
                $uri = (Get-MSCloudLoginConnectionProfile -Workload MicrosoftGraph).ResourceUrl + "beta/roleManagement/defender/roleDefinitions/$Id"
                [array]$definition = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            }
            if ($definition.Length -eq 0)
            {
                Write-Verbose -Message "No Defender Role Definition {$Id} was found by Identity. Trying to retrieve by DisplayName"
                $uri = (Get-MSCloudLoginConnectionProfile -Workload MicrosoftGraph).ResourceUrl + "beta/roleManagement/defender/roleDefinitions/?`$filter=displayName eq '$DisplayName'"
                [Array]$definition = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
            }

            if ($definition.Length -gt 1)
            {
                throw "Multiple definitions with display name {$DisplayName} were found. Please ensure only one instance exists."
            }

            if ($null -eq $definition)
            {
                Write-Verbose -Message "No Defender Role Definition {$DisplayName} was found by Display Name. Instance doesn't exist."
                return $nullResult
            }
        }
        else
        {
            $definition = $Script:exportedInstance
        }

        $rolePermissionsValue = $null
        if ($definition.RolePermissions.Length -gt 0)
        {
            $rolePermissionsValue = @(
                @{
                        allowedResourceActions = $definition.RolePermissions.allowedResourceActions
                }
            )
        }

        return @{
            Id                    = $definition.Id
            DisplayName           = $definition.DisplayName
            Description           = $definition.Description
            RolePermissions       = $rolePermissionsValue
            Ensure                = 'Present'
            Credential            = $Credential
            ApplicationId         = $ApplicationId
            ApplicationSecret     = $ApplicationSecret
            TenantId              = $TenantId
            CertificateThumbprint = $CertificateThumbprint
            ManagedIdentity       = $ManagedIdentity.IsPresent
            AccessTokens          = $AccessTokens
        }
    }
    catch
    {
        New-M365DSCLogEntry -Message 'Error retrieving data:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        throw
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Id,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $RolePermissions,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $TenantId,

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

    Write-Verbose -Message "Setting configuration of the Defender Role Definition with DisplayName {$DisplayName}"

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName -replace 'MSFT_', ''
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $currentPolicy = Get-TargetResource @PSBoundParameters
    $Identity = $currentPolicy.Identity

    if ($Ensure -eq 'Present' -and $currentPolicy.Ensure -eq 'Absent')
    {
        Write-Verbose -Message "Creating new iOS App Protection Policy {$DisplayName}"
        $createParameters = Remove-M365DSCAuthenticationParameter -BoundParameters $PSBoundParameters
        $createParameters.Remove('Identity')
        $createParameters.Remove('Assignments')
        $createParameters.Remove('Apps')
        $createParameters.Remove('DeployedAppCount')
        $createParameters.TargetedAppManagementLevels = $createParameters.TargetedAppManagementLevels -join ','

        $myExemptedAppProtocols = @()
        foreach ($exemptedAppProtocol in $ExemptedAppProtocols)
        {
            $myExemptedAppProtocols += @{
                Name  = $exemptedAppProtocol.Split(':')[0]
                Value = $exemptedAppProtocol.Split(':')[1]
            }
        }
        $createParameters.ExemptedAppProtocols = $myExemptedAppProtocols

        # Remove empty string parameters that the cmdlet can't handle
        $arrayTemp = @("MinimumWarningSdkVersion", "MaximumRequiredOsVersion", "MaximumWarningOsVersion", "MaximumWipeOsVersion")
        foreach ($item in $arrayTemp)
        {
            if ([System.String]::IsNullOrEmpty($createParameters.$item))
            {
                $createParameters.Remove($item)
            }
        }

        $policy = New-MgBetaDeviceAppManagementiOSManagedAppProtection -BodyParameter $createParameters
        if ($policy.Id)
        {
            Write-Verbose -Message "Update targetApps for iOS App Protection Policy with Id {$($policy.Id)} and DisplayName {$DisplayName}"
            $targetApps = Get-IntuneAppProtectionPolicyiOSAppsToHashtable -Apps $Apps -AppGroupType $AppGroupType
            $Url = (Get-MSCloudLoginConnectionProfile -Workload MicrosoftGraph).ResourceUrl + "beta/deviceAppManagement/iosManagedAppProtections('$($policy.Id)')/targetApps"
            Invoke-MgGraphRequest -Method POST -Uri $Url -Body $targetApps

            $assignmentsHash = ConvertTo-IntunePolicyAssignment -IncludeDeviceFilter:$true -Assignments $Assignments
            Update-DeviceConfigurationPolicyAssignment `
                -DeviceConfigurationPolicyId $policy.Id `
                -Targets $assignmentsHash `
                -Repository 'deviceAppManagement/iosManagedAppProtections'
        }
    }
    elseif ($Ensure -eq 'Present' -and $currentPolicy.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Updating existing iOS App Protection Policy {$DisplayName}"
        $updateParameters = Remove-M365DSCAuthenticationParameter -BoundParameters $PSBoundParameters
        $updateParameters.Remove('Identity')
        $updateParameters.Remove('Assignments')
        $updateParameters.Remove('Apps')
        $updateParameters.Remove('DeployedAppCount')
        $updateParameters.TargetedAppManagementLevels = $updateParameters.TargetedAppManagementLevels -join ','

        # Remove empty string parameters that the cmdlet can't handle
        $arrayTemp = @("MinimumWarningSdkVersion", "MaximumRequiredOsVersion", "MaximumWarningOsVersion", "MaximumWipeOsVersion")
        foreach ($item in $arrayTemp)
        {
            if ([System.String]::IsNullOrEmpty($updateParameters.$item))
            {
                $updateParameters.Remove($item)
            }
        }

        $myExemptedAppProtocols = @()
        foreach ($exemptedAppProtocol in $ExemptedAppProtocols)
        {
            $myExemptedAppProtocols += @{
                Name  = $exemptedAppProtocol.Split(':')[0]
                Value = $exemptedAppProtocol.Split(':')[1]
            }
        }
        $updateParameters.ExemptedAppProtocols = $myExemptedAppProtocols
        Update-MgBetaDeviceAppManagementiOSManagedAppProtection -IosManagedAppProtectionId $Identity -BodyParameter $updateParameters

        Write-Verbose -Message "Updating targetApps for iOS App Protection Policy with Id {$Identity} and DisplayName {$DisplayName}"
        $targetApps = Get-IntuneAppProtectionPolicyiOSAppsToHashtable -Apps $Apps -AppGroupType $AppGroupType
        $Url = (Get-MSCloudLoginConnectionProfile -Workload MicrosoftGraph).ResourceUrl + "beta/deviceAppManagement/iosManagedAppProtections('$($Identity)')/targetApps"
        Invoke-MgGraphRequest -Method POST -Uri $Url -Body $targetApps

        $assignmentsHash = ConvertTo-IntunePolicyAssignment -IncludeDeviceFilter:$true -Assignments $Assignments
        Update-DeviceConfigurationPolicyAssignment `
            -DeviceConfigurationPolicyId $Identity `
            -Targets $assignmentsHash `
            -Repository 'deviceAppManagement/iosManagedAppProtections'

    }
    elseif ($Ensure -eq 'Absent' -and $currentPolicy.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Removing iOS App Protection Policy {$DisplayName}"
        Remove-MgBetaDeviceAppManagementiOSManagedAppProtection -IosManagedAppProtectionId $Identity
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Id,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $RolePermissions,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $TenantId,

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

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $result = Test-M365DSCTargetResource -DesiredValues $PSBoundParameters `
                                         -ResourceName $($MyInvocation.MyCommand.Source).Replace('MSFT_', '')
    return $result
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

    #Ensure the proper dependencies are installed in the current environment
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName -replace 'MSFT_', ''
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    try
    {
        $uri = (Get-MSCloudLoginConnectionProfile -Workload MicrosoftGraph).ResourceUrl + "beta/roleManagement/defender/roleDefinitions"
        [array]$roleDefinitions = (Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop).value

        $i = 1
        $dscContent = ''
        if ($roleDefinitions.Length -eq 0)
        {
            Write-M365DSCHost -Message $Global:M365DSCEmojiGreenCheckMark -CommitWrite
        }
        else
        {
            Write-M365DSCHost -Message "`r`n" -DeferWrite
        }
        foreach ($role in $roleDefinitions)
        {
            if ($null -ne $Global:M365DSCExportResourceInstancesCount)
            {
                $Global:M365DSCExportResourceInstancesCount++
            }

            Write-M365DSCHost -Message "    |---[$i/$($roleDefinitions.Count)] $($role.displayName)" -DeferWrite
            $params = @{
                Id                    = $role.id
                DisplayName           = $role.displayName
                Ensure                = 'Present'
                Credential            = $Credential
                ApplicationID         = $ApplicationId
                TenantId              = $TenantId
                ApplicationSecret     = $ApplicationSecret
                CertificateThumbprint = $CertificateThumbprint
                ManagedIdentity       = $ManagedIdentity.IsPresent
                AccessTokens          = $AccessTokens
            }

            $Script:exportedInstance = $role
            $Results = Get-TargetResource @Params

            if ($Results.RolePermissions)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString -ComplexObject $Results.RolePermissions `
                                                                             -CIMInstanceName 'DefenderRoleDefinitionRolePermissions'
                if ($complexTypeStringResult)
                {
                    $Results.RolePermissions = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('RolePermissions') | Out-Null
                }
            }

            $currentDSCBlock = Get-M365DSCExportContentForResource -ResourceName $ResourceName `
                -ConnectionMode $ConnectionMode `
                -ModulePath $PSScriptRoot `
                -Results $Results `
                -Credential $Credential `
                -NoEscape @('RolePermissions')
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
        New-M365DSCLogEntry -Message 'Error during Export:' `
                -Exception $_ `
                -Source $($MyInvocation.MyCommand.Source) `
                -TenantId $TenantId `
                -Credential $Credential

        throw
    }
}

Export-ModuleMember -Function *-TargetResource

Confirm-M365DSCModuleDependency -ModuleName 'MSFT_IntuneDeviceConfigurationPolicyAndroidWorkProfile'

function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $True)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockFaceUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockFingerprintUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockIrisUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockTrustAgents,

        [Parameter()]
        [System.Int32]
        $PasswordExpirationDays,

        [Parameter()]
        [System.Int32]
        $PasswordMinimumLength,

        [Parameter()]
        [System.Int32]
        $PasswordMinutesOfInactivityBeforeScreenTimeout,

        [Parameter()]
        [System.Int32]
        $PasswordPreviousPasswordBlockCount,

        [Parameter()]
        [System.Int32]
        $PasswordSignInFailureCountBeforeFactoryReset,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'lowSecurityBiometric', 'required', 'atLeastNumeric', 'numericComplex', 'atLeastAlphabetic', 'atLeastAlphanumeric', 'alphanumericWithSymbols')]
        $PasswordRequiredType,

        [Parameter()]
        [System.String]
        [ValidateSet('none', 'low', 'medium', 'high')]
        $RequiredPasswordComplexity,

        [Parameter()]
        [System.Boolean]
        $WorkProfileAllowAppInstallsFromUnknownSources,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'preventAny', 'allowPersonalToWork', 'noRestrictions')]
        $WorkProfileDataSharingType,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockNotificationsWhileDeviceLocked,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockAddingAccounts,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBluetoothEnableContactSharing,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockScreenCapture,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileCallerId,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCamera,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileContactsSearch,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileCopyPaste,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'prompt', 'autoGrant', 'autoDeny')]
        $WorkProfileDefaultAppPermissionPolicy,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockFaceUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockFingerprintUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockIrisUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockTrustAgents,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordExpirationDays,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinimumLength,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinNumericCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinNonLetterCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinLetterCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinLowerCaseCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinUpperCaseCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinSymbolCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinutesOfInactivityBeforeScreenTimeout,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordPreviousPasswordBlockCount,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordSignInFailureCountBeforeFactoryReset,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'lowSecurityBiometric', 'required', 'atLeastNumeric', 'numericComplex', 'atLeastAlphabetic', 'atLeastAlphanumeric', 'alphanumericWithSymbols')]
        $WorkProfilePasswordRequiredType,

        [Parameter()]
        [System.String]
        [ValidateSet('none', 'low', 'medium', 'high')]
        $WorkProfileRequiredPasswordComplexity,

        [Parameter()]
        [System.Boolean]
        $WorkProfileRequirePassword,

        [Parameter()]
        [System.Boolean]
        $SecurityRequireVerifyApps,

        [Parameter()]
        [System.String]
        $VpnAlwaysOnPackageIdentifier,

        [Parameter()]
        [System.Boolean]
        $VpnEnableAlwaysOnLockdownMode,

        [Parameter()]
        [System.Boolean]
        $WorkProfileAllowWidgets,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockPersonalAppInstallsFromUnknownSources,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,

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

    Write-Verbose -Message "Getting configuration of the Intune Device Configuration Policy Android for Work Profile {$DisplayName}"

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

            $policy = Get-MgBetaDeviceManagementDeviceConfiguration -All -Filter "DisplayName eq '$($DisplayName -replace "'", "''")'" `
                -ErrorAction Stop | Where-Object -FilterScript { $_.'@odata.type' -eq '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration' }

            if ($null -eq $policy)
            {
                Write-Verbose -Message "No Intune Device Configuration Policy Android for Work Profile with {$DisplayName} was found"
                return $nullResult
            }
        }
        else
        {
            $policy = $Script:exportedInstance
        }

        Write-Verbose -Message "An Intune Device Configuration Policy Android for Work Profile with {$DisplayName} was found"
        $results = @{
            Description                                               = $policy.Description
            DisplayName                                               = $policy.DisplayName
            RoleScopeTagIds                                           = $policy.RoleScopeTagIds
            PasswordBlockFaceUnlock                                   = $policy.passwordBlockFaceUnlock
            PasswordBlockFingerprintUnlock                            = $policy.passwordBlockFingerprintUnlock
            PasswordBlockIrisUnlock                                   = $policy.passwordBlockIrisUnlock
            PasswordBlockTrustAgents                                  = $policy.passwordBlockTrustAgents
            PasswordExpirationDays                                    = $policy.passwordExpirationDays
            PasswordMinimumLength                                     = $policy.passwordMinimumLength
            PasswordMinutesOfInactivityBeforeScreenTimeout            = $policy.passwordMinutesOfInactivityBeforeScreenTimeout
            PasswordPreviousPasswordBlockCount                        = $policy.passwordPreviousPasswordBlockCount
            PasswordSignInFailureCountBeforeFactoryReset              = $policy.passwordSignInFailureCountBeforeFactoryReset
            PasswordRequiredType                                      = $policy.passwordRequiredType
            RequiredPasswordComplexity                                = $policy.requiredPasswordComplexity
            WorkProfileAllowAppInstallsFromUnknownSources             = $policy.workProfileAllowAppInstallsFromUnknownSources
            WorkProfileDataSharingType                                = $policy.workProfileDataSharingType
            WorkProfileBlockNotificationsWhileDeviceLocked            = $policy.workProfileBlockNotificationsWhileDeviceLocked
            WorkProfileBlockAddingAccounts                            = $policy.workProfileBlockAddingAccounts
            WorkProfileBluetoothEnableContactSharing                  = $policy.workProfileBluetoothEnableContactSharing
            WorkProfileBlockScreenCapture                             = $policy.workProfileBlockScreenCapture
            WorkProfileBlockCrossProfileCallerId                      = $policy.workProfileBlockCrossProfileCallerId
            WorkProfileBlockCamera                                    = $policy.workProfileBlockCamera
            WorkProfileBlockCrossProfileContactsSearch                = $policy.workProfileBlockCrossProfileContactsSearch
            WorkProfileBlockCrossProfileCopyPaste                     = $policy.workProfileBlockCrossProfileCopyPaste
            WorkProfileDefaultAppPermissionPolicy                     = $policy.workProfileDefaultAppPermissionPolicy
            WorkProfilePasswordBlockFaceUnlock                        = $policy.workProfilePasswordBlockFaceUnlock
            WorkProfilePasswordBlockFingerprintUnlock                 = $policy.workProfilePasswordBlockFingerprintUnlock
            WorkProfilePasswordBlockIrisUnlock                        = $policy.workProfilePasswordBlockIrisUnlock
            WorkProfilePasswordBlockTrustAgents                       = $policy.workProfilePasswordBlockTrustAgents
            WorkProfilePasswordExpirationDays                         = $policy.workProfilePasswordExpirationDays
            WorkProfilePasswordMinimumLength                          = $policy.workProfilePasswordMinimumLength
            WorkProfilePasswordMinNumericCharacters                   = $policy.workProfilePasswordMinNumericCharacters
            WorkProfilePasswordMinNonLetterCharacters                 = $policy.workProfilePasswordMinNonLetterCharacters
            WorkProfilePasswordMinLetterCharacters                    = $policy.workProfilePasswordMinLetterCharacters
            WorkProfilePasswordMinLowerCaseCharacters                 = $policy.workProfilePasswordMinLowerCaseCharacters
            WorkProfilePasswordMinUpperCaseCharacters                 = $policy.workProfilePasswordMinUpperCaseCharacters
            WorkProfilePasswordMinSymbolCharacters                    = $policy.workProfilePasswordMinSymbolCharacters
            WorkProfilePasswordMinutesOfInactivityBeforeScreenTimeout = $policy.workProfilePasswordMinutesOfInactivityBeforeScreenTimeout
            WorkProfilePasswordPreviousPasswordBlockCount             = $policy.workProfilePasswordPreviousPasswordBlockCount
            WorkProfilePasswordSignInFailureCountBeforeFactoryReset   = $policy.workProfilePasswordSignInFailureCountBeforeFactoryReset
            WorkProfilePasswordRequiredType                           = $policy.workProfilePasswordRequiredType
            WorkProfileRequiredPasswordComplexity                     = $policy.workProfileRequiredPasswordComplexity
            WorkProfileRequirePassword                                = $policy.workProfileRequirePassword
            SecurityRequireVerifyApps                                 = $policy.securityRequireVerifyApps
            VpnAlwaysOnPackageIdentifier                              = $policy.vpnAlwaysOnPackageIdentifier
            VpnEnableAlwaysOnLockdownMode                             = $policy.vpnEnableAlwaysOnLockdownMode
            WorkProfileAllowWidgets                                   = $policy.workProfileAllowWidgets
            WorkProfileBlockPersonalAppInstallsFromUnknownSources     = $policy.workProfileBlockPersonalAppInstallsFromUnknownSources
            Ensure                                                    = 'Present'
            Credential                                                = $Credential
            ApplicationId                                             = $ApplicationId
            TenantId                                                  = $TenantId
            ApplicationSecret                                         = $ApplicationSecret
            CertificateThumbprint                                     = $CertificateThumbprint
            ManagedIdentity                                           = $ManagedIdentity.IsPresent
            AccessTokens                                              = $AccessTokens
        }

        $assignmentsValues = Get-MgBetaDeviceManagementDeviceConfigurationAssignment -DeviceConfigurationId $policy.Id
        $assignmentResult = @()
        if ($assignmentsValues.Count -gt 0)
        {
            $assignmentResult += ConvertFrom-IntunePolicyAssignment `
                -IncludeDeviceFilter:$true `
                -Assignments ($assignmentsValues)
        }
        $results.Add('Assignments', $assignmentResult)

        return $results
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
        [Parameter(Mandatory = $True)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockFaceUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockFingerprintUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockIrisUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockTrustAgents,

        [Parameter()]
        [System.Int32]
        $PasswordExpirationDays,

        [Parameter()]
        [System.Int32]
        $PasswordMinimumLength,

        [Parameter()]
        [System.Int32]
        $PasswordMinutesOfInactivityBeforeScreenTimeout,

        [Parameter()]
        [System.Int32]
        $PasswordPreviousPasswordBlockCount,

        [Parameter()]
        [System.Int32]
        $PasswordSignInFailureCountBeforeFactoryReset,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'lowSecurityBiometric', 'required', 'atLeastNumeric', 'numericComplex', 'atLeastAlphabetic', 'atLeastAlphanumeric', 'alphanumericWithSymbols')]
        $PasswordRequiredType,

        [Parameter()]
        [System.String]
        [ValidateSet('none', 'low', 'medium', 'high')]
        $RequiredPasswordComplexity,

        [Parameter()]
        [System.Boolean]
        $WorkProfileAllowAppInstallsFromUnknownSources,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'preventAny', 'allowPersonalToWork', 'noRestrictions')]
        $WorkProfileDataSharingType,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockNotificationsWhileDeviceLocked,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockAddingAccounts,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBluetoothEnableContactSharing,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockScreenCapture,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileCallerId,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCamera,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileContactsSearch,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileCopyPaste,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'prompt', 'autoGrant', 'autoDeny')]
        $WorkProfileDefaultAppPermissionPolicy,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockFaceUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockFingerprintUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockIrisUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockTrustAgents,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordExpirationDays,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinimumLength,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinNumericCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinNonLetterCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinLetterCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinLowerCaseCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinUpperCaseCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinSymbolCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinutesOfInactivityBeforeScreenTimeout,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordPreviousPasswordBlockCount,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordSignInFailureCountBeforeFactoryReset,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'lowSecurityBiometric', 'required', 'atLeastNumeric', 'numericComplex', 'atLeastAlphabetic', 'atLeastAlphanumeric', 'alphanumericWithSymbols')]
        $WorkProfilePasswordRequiredType,

        [Parameter()]
        [System.String]
        [ValidateSet('none', 'low', 'medium', 'high')]
        $WorkProfileRequiredPasswordComplexity,

        [Parameter()]
        [System.Boolean]
        $WorkProfileRequirePassword,

        [Parameter()]
        [System.Boolean]
        $SecurityRequireVerifyApps,

        [Parameter()]
        [System.String]
        $VpnAlwaysOnPackageIdentifier,

        [Parameter()]
        [System.Boolean]
        $VpnEnableAlwaysOnLockdownMode,

        [Parameter()]
        [System.Boolean]
        $WorkProfileAllowWidgets,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockPersonalAppInstallsFromUnknownSources,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,

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

    $currentPolicy = Get-TargetResource @PSBoundParameters
    $boundParameters = Remove-M365DSCAuthenticationParameter -BoundParameters $PSBoundParameters

    if ($Ensure -eq 'Present' -and $currentPolicy.Ensure -eq 'Absent')
    {
        Write-Verbose -Message "Creating new Device Configuration Policy {$DisplayName}"
        $boundParameters.Remove('DisplayName') | Out-Null
        $boundParameters.Remove('Description') | Out-Null
        $boundParameters.Remove('Assignments') | Out-Null

        $AdditionalProperties = Get-M365DSCIntuneDeviceConfigurationPolicyAndroidWorkProfileAdditionalProperties -Properties $boundParameters
        $policy = New-MgBetaDeviceManagementDeviceConfiguration -DisplayName $DisplayName `
            -Description $Description `
            -AdditionalProperties $AdditionalProperties

        #region Assignments
        $assignmentsHash = ConvertTo-IntunePolicyAssignment -IncludeDeviceFilter:$true -Assignments $Assignments
        if ($policy.id)
        {
            Update-DeviceConfigurationPolicyAssignment -DeviceConfigurationPolicyId $policy.id `
                -Targets $assignmentsHash `
                -Repository 'deviceManagement/deviceConfigurations'
        }
        #endregion
    }
    elseif ($Ensure -eq 'Present' -and $currentPolicy.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Updating existing Device Configuration Policy {$DisplayName}"
        $configDevicePolicy = Get-MgBetaDeviceManagementDeviceConfiguration -Filter "DisplayName eq '$($Displayname -replace "'", "''")'" -ErrorAction SilentlyContinue | Where-Object `
            -FilterScript {
            $_.'@odata.type' -eq '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration'
        }

        $boundParameters.Remove('DisplayName') | Out-Null
        $boundParameters.Remove('Description') | Out-Null
        $boundParameters.Remove('Assignments') | Out-Null

        $AdditionalProperties = Get-M365DSCIntuneDeviceConfigurationPolicyAndroidWorkProfileAdditionalProperties -Properties $boundParameters
        Update-MgBetaDeviceManagementDeviceConfiguration -AdditionalProperties $AdditionalProperties `
            -Description $Description `
            -DeviceConfigurationId $configDevicePolicy.Id

        #region Assignments
        $assignmentsHash = ConvertTo-IntunePolicyAssignment -IncludeDeviceFilter:$true -Assignments $Assignments
        Update-DeviceConfigurationPolicyAssignment `
            -DeviceConfigurationPolicyId $configDevicePolicy.id `
            -Targets $assignmentsHash `
            -Repository 'deviceManagement/deviceConfigurations'
        #endregion
    }
    elseif ($Ensure -eq 'Absent' -and $currentPolicy.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Removing Device Configuration Policy {$DisplayName}"
        $configDevicePolicy = Get-MgBetaDeviceManagementDeviceConfiguration -Filter "DisplayName eq '$($Displayname -replace "'", "''")'" -ErrorAction SilentlyContinue | Where-Object `
            -FilterScript {
            $_.'@odata.type' -eq '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration' `
        }

        Remove-MgBetaDeviceManagementDeviceConfiguration -DeviceConfigurationId $configDevicePolicy.Id
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $True)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockFaceUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockFingerprintUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockIrisUnlock,

        [Parameter()]
        [System.Boolean]
        $PasswordBlockTrustAgents,

        [Parameter()]
        [System.Int32]
        $PasswordExpirationDays,

        [Parameter()]
        [System.Int32]
        $PasswordMinimumLength,

        [Parameter()]
        [System.Int32]
        $PasswordMinutesOfInactivityBeforeScreenTimeout,

        [Parameter()]
        [System.Int32]
        $PasswordPreviousPasswordBlockCount,

        [Parameter()]
        [System.Int32]
        $PasswordSignInFailureCountBeforeFactoryReset,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'lowSecurityBiometric', 'required', 'atLeastNumeric', 'numericComplex', 'atLeastAlphabetic', 'atLeastAlphanumeric', 'alphanumericWithSymbols')]
        $PasswordRequiredType,

        [Parameter()]
        [System.String]
        [ValidateSet('none', 'low', 'medium', 'high')]
        $RequiredPasswordComplexity,

        [Parameter()]
        [System.Boolean]
        $WorkProfileAllowAppInstallsFromUnknownSources,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'preventAny', 'allowPersonalToWork', 'noRestrictions')]
        $WorkProfileDataSharingType,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockNotificationsWhileDeviceLocked,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockAddingAccounts,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBluetoothEnableContactSharing,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockScreenCapture,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileCallerId,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCamera,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileContactsSearch,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockCrossProfileCopyPaste,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'prompt', 'autoGrant', 'autoDeny')]
        $WorkProfileDefaultAppPermissionPolicy,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockFaceUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockFingerprintUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockIrisUnlock,

        [Parameter()]
        [System.Boolean]
        $WorkProfilePasswordBlockTrustAgents,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordExpirationDays,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinimumLength,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinNumericCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinNonLetterCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinLetterCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinLowerCaseCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinUpperCaseCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinSymbolCharacters,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordMinutesOfInactivityBeforeScreenTimeout,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordPreviousPasswordBlockCount,

        [Parameter()]
        [System.Int32]
        $WorkProfilePasswordSignInFailureCountBeforeFactoryReset,

        [Parameter()]
        [System.String]
        [ValidateSet('deviceDefault', 'lowSecurityBiometric', 'required', 'atLeastNumeric', 'numericComplex', 'atLeastAlphabetic', 'atLeastAlphanumeric', 'alphanumericWithSymbols')]
        $WorkProfilePasswordRequiredType,

        [Parameter()]
        [System.String]
        [ValidateSet('none', 'low', 'medium', 'high')]
        $WorkProfileRequiredPasswordComplexity,

        [Parameter()]
        [System.Boolean]
        $WorkProfileRequirePassword,

        [Parameter()]
        [System.Boolean]
        $SecurityRequireVerifyApps,

        [Parameter()]
        [System.String]
        $VpnAlwaysOnPackageIdentifier,

        [Parameter()]
        [System.Boolean]
        $VpnEnableAlwaysOnLockdownMode,

        [Parameter()]
        [System.Boolean]
        $WorkProfileAllowWidgets,

        [Parameter()]
        [System.Boolean]
        $WorkProfileBlockPersonalAppInstallsFromUnknownSources,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,

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
        [array]$policies = Get-MgBetaDeviceManagementDeviceConfiguration `
            -ErrorAction Stop -All:$true -Filter $Filter | Where-Object `
            -FilterScript { $_.'@odata.type' -eq '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration' }
        $i = 1
        $dscContent = ''
        if ($policies.Length -eq 0)
        {
            Write-M365DSCHost -Message $Global:M365DSCEmojiGreenCheckMark -CommitWrite
        }
        else
        {
            Write-M365DSCHost -Message "`r`n" -DeferWrite
        }
        foreach ($policy in $policies)
        {
            if ($null -ne $Global:M365DSCExportResourceInstancesCount)
            {
                $Global:M365DSCExportResourceInstancesCount++
            }

            Write-M365DSCHost -Message "    |---[$i/$($policies.Count)] $($policy.DisplayName)" -DeferWrite
            $params = @{
                DisplayName           = $policy.DisplayName
                Ensure                = 'Present'
                Credential            = $Credential
                ApplicationId         = $ApplicationId
                TenantId              = $TenantId
                ApplicationSecret     = $ApplicationSecret
                CertificateThumbprint = $CertificateThumbprint
                ManagedIdentity       = $ManagedIdentity.IsPresent
                AccessTokens          = $AccessTokens
            }

            $Script:exportedInstance = $policy
            $Results = Get-TargetResource @Params
            if ($Results.Assignments)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString -ComplexObject ([Array]$Results.Assignments) -CIMInstanceName DeviceManagementConfigurationPolicyAssignments

                if ($complexTypeStringResult)
                {
                    $Results.Assignments = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('Assignments') | Out-Null
                }
            }


            $currentDSCBlock = Get-M365DSCExportContentForResource -ResourceName $ResourceName `
                -ConnectionMode $ConnectionMode `
                -ModulePath $PSScriptRoot `
                -Results $Results `
                -Credential $Credential `
                -NoEscape @('Assignments')

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
        if ($_.Exception -like '*401*' -or $_.ErrorDetails.Message -like "*`"ErrorCode`":`"Forbidden`"*" -or `
                $_.Exception -like '*Request not applicable to target tenant*')
        {
            Write-M365DSCHost -Message "`r`n    $($Global:M365DSCEmojiYellowCircle) The current tenant is not registered for Intune."
        }
        else
        {
            New-M365DSCLogEntry -Message 'Error during Export:' `
                -Exception $_ `
                -Source $($MyInvocation.MyCommand.Source) `
                -TenantId $TenantId `
                -Credential $Credential

            throw
        }
    }
}

function Get-M365DSCIntuneDeviceConfigurationPolicyAndroidWorkProfileAdditionalProperties
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = 'true')]
        [System.Collections.Hashtable]
        $Properties
    )

    $results = @{'@odata.type' = '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration' }
    foreach ($property in $properties.Keys)
    {
        if ($property -ne 'Verbose')
        {
            $propertyName = $property[0].ToString().ToLower() + $property.Substring(1, $property.Length - 1)
            $propertyValue = $properties.$property
            $results.Add($propertyName, $propertyValue)
        }
    }
    return $results
}

Export-ModuleMember -Function *-TargetResource

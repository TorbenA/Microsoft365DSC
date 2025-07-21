function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        #region resource generator code
        [Parameter()]
        [System.String]
        $Id,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String]
        $Developer,

        [Parameter()]
        [System.String]
        $InformationUrl,

        [Parameter()]
        [System.Boolean]
        $IsFeatured,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $LargeIcon,

        [Parameter()]
        [System.String]
        $Notes,

        [Parameter()]
        [System.String]
        $Owner,

        [Parameter()]
        [System.String]
        $PrivacyInformationUrl,

        [Parameter()]
        [System.String]
        $Publisher,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Categories,

        [Parameter()]
        [System.String]
        $InstallCommandLine,

        [Parameter()]
        [System.String]
        $UninstallCommandLine,

        [Parameter()]
        [System.String[]]
        $AllowedArchitectures,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $MinimumSupportedOperatingSystem,

        [Parameter()]
        [System.Int32]
        $MinimumFreeDiskSpaceInMB,

        [Parameter()]
        [System.Int32]
        $MinimumMemoryInMB,

        [Parameter()]
        [System.Int32]
        $MinimumNumberOfProcessors,

        [Parameter()]
        [System.Int32]
        $MinimumCpuSpeedInMHz,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $DetectionRules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $RequirementRules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Rules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $InstallExperience,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $ReturnCodes,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $MsiInformation,

        [Parameter()]
        [System.String]
        $SetupFilePath,

        [Parameter()]
        [System.String]
        $MinimumSupportedWindowsRelease,

        [Parameter()]
        [System.String]
        $DisplayVersion,

        [Parameter()]
        [System.Boolean]
        $AllowAvailableUninstall,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,
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

    Write-Verbose -Message "Getting configuration for the Intune Mobile Apps Win32 App for Windows10 with Id {$Id} and DisplayName {$DisplayName}"

    try
    {
        if (-not $Script:exportedInstance -or $Script:exportedInstance.DisplayName -ne $DisplayName)
        {

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

            $nullResult = $PSBoundParameters
            $nullResult.Ensure = 'Absent'

            $getValue = $null

            #region resource generator code
            if (-not [System.String]::IsNullOrEmpty($Id))
            {
                $getValue = Get-MgBetaDeviceAppManagementMobileApp -MobileAppId $Id -ExpandProperty 'Categories' -ErrorAction SilentlyContinue
            }

            if ($null -eq $getValue)
            {
                Write-Verbose -Message "Could not find an Intune Mobile Apps Win32 App for Windows10 with Id {$Id}"

                if (-not [System.String]::IsNullOrEmpty($DisplayName))
                {
                    $getValue = Get-MgBetaDeviceAppManagementMobileApp `
                        -Filter "DisplayName eq '$($DisplayName -replace "'", "''")' and isof('microsoft.graph.win32LobApp')" `
                        -ExpandProperty 'Categories' `
                        -All `
                        -ErrorAction SilentlyContinue
                }
            }
            #endregion
            if ($null -eq $getValue)
            {
                Write-Verbose -Message "Could not find an Intune Mobile Apps Win32 App for Windows10 with DisplayName {$DisplayName}."
                return $nullResult
            }
        }
        else
        {
            $getValue = Get-MgBetaDeviceAppManagementMobileApp -MobileAppId $getValue.Id -ExpandProperty 'Categories'
        }
        $Id = $getValue.Id
        Write-Verbose -Message "An Intune Mobile Apps Win32 App for Windows10 with Id {$Id} and DisplayName {$DisplayName} was found"

        #region resource generator code
        $complexCategories = @()
        foreach ($category in $getValue.Categories)
        {
            $myCategory = @{}
            $myCategory.Add('Id', $category.id)
            $myCategory.Add('DisplayName', $category.displayName)
            $complexCategories += $myCategory
        }
        $complexLargeIcon = $null
        if ($null -ne $getValue.LargeIcon.Value)
        {
            $complexLargeIcon = @{}
            $complexLargeIcon.Add('Type', $getValue.LargeIcon.Type)
            $complexLargeIcon.Add('Value', [System.Convert]::ToBase64String($getValue.LargeIcon.Value))
        }

        $complexMinimumSupportedOperatingSystem = [ordered]@{}
        $complexMinimumSupportedOperatingSystem.Add('V8_0', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v8_0)
        $complexMinimumSupportedOperatingSystem.Add('V8_1', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v8_1)
        $complexMinimumSupportedOperatingSystem.Add('V10_0', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_0)
        $complexMinimumSupportedOperatingSystem.Add('V10_1607', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_1607)
        $complexMinimumSupportedOperatingSystem.Add('V10_1703', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_1703)
        $complexMinimumSupportedOperatingSystem.Add('V10_1709', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_1709)
        $complexMinimumSupportedOperatingSystem.Add('V10_1803', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_1803)
        $complexMinimumSupportedOperatingSystem.Add('V10_1809', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_1809)
        $complexMinimumSupportedOperatingSystem.Add('V10_1903', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_1903)
        $complexMinimumSupportedOperatingSystem.Add('V10_1909', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_1909)
        $complexMinimumSupportedOperatingSystem.Add('V10_2004', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_2004)
        $complexMinimumSupportedOperatingSystem.Add('V10_2H20', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_2H20)
        $complexMinimumSupportedOperatingSystem.Add('V10_21H1', $getValue.AdditionalProperties.minimumSupportedOperatingSystem.v10_21H1)
        if ($complexMinimumSupportedOperatingSystem.Values.Where({ $null -ne $_ }).Count -eq 0)
        {
            $complexMinimumSupportedOperatingSystem = $null
        }

        $complexDetectionRules = @()
        foreach ($detectionRule in $getValue.AdditionalProperties.detectionRules)
        {
            $detectionType = $detectionRule.'@odata.type'.Replace('#microsoft.graph.win32LobApp', '').Replace('Detection', '')
            $baseDetectionRule = @{
                OdataType      = $detectionType
                Operator       = $detectionRule.operator
                DetectionValue = $detectionRule.detectionValue
            }
            switch ($detectionType)
            {
                'FileSystem' {
                    $baseDetectionRule.Add('Check32BitOn64System', $detectionRule.check32BitOn64System)
                    $baseDetectionRule.Add('Path', $detectionRule.path)
                    $baseDetectionRule.Add('FileOrFolderName', $detectionRule.fileOrFolderName)
                    $baseDetectionRule.Add('FileSystemDetectionType', $detectionRule.detectionType)
                }
                'Registry' {
                    $baseDetectionRule.Add('Check32BitOn64System', $detectionRule.check32BitOn64System)
                    $baseDetectionRule.Add('KeyPath', $detectionRule.keyPath)
                    $baseDetectionRule.Add('RegistryDetectionType', $detectionRule.detectionType)
                    $baseDetectionRule.Add('ValueName', $detectionRule.registryValueName)
                }
                "ProductCode" {
                    $baseDetectionRule.Add('ProductCode', $detectionRule.productCode)
                    $baseDetectionRule.Add('ProductVersionOperator', $detectionRule.productVersionOperator)
                    $baseDetectionRule.Add('ProductVersion', $detectionRule.productVersion)
                }
                "PowerShellScript" {
                    $baseDetectionRule.Add('Script', [System.Convert]::FromBase64String($detectionRule.scriptContent))
                    $baseDetectionRule.Add('RunAs32Bit', $detectionRule.runAs32Bit)
                    $baseDetectionRule.Add('EnforceSignatureCheck', $detectionRule.enforceSignatureCheck)
                }
            }
            $complexDetectionRules += $baseDetectionRule
        }

        $complexRequirementRules = @()
        foreach ($requirementRule in $getValue.AdditionalProperties.requirementRules)
        {
            $requirementType = $requirementRule.'@odata.type'.Replace('#microsoft.graph.win32LobApp', '').Replace('Requirement', '')
            $baseRequirementRule = @{
                OdataType      = $requirementType
                Operator       = $requirementRule.operator
                DetectionValue = $requirementRule.detectionValue
            }
            switch ($requirementType)
            {
                'FileSystem' {
                    $baseRequirementRule.Add('Check32BitOn64System', $requirementRule.check32BitOn64System)
                    $baseRequirementRule.Add('Path', $requirementRule.path)
                    $baseRequirementRule.Add('FileOrFolderName', $requirementRule.fileOrFolderName)
                    $baseRequirementRule.Add('FileSystemDetectionType', $requirementRule.detectionType)
                }
                'Registry' {
                    $baseRequirementRule.Add('Check32BitOn64System', $requirementRule.check32BitOn64System)
                    $baseRequirementRule.Add('KeyPath', $requirementRule.keyPath)
                    $baseRequirementRule.Add('RegistryDetectionType', $requirementRule.detectionType)
                    $baseRequirementRule.Add('ValueName', $requirementRule.registryValueName)
                }
                "PowerShellScript" {
                    $baseRequirementRule.Add('Script', [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($requirementRule.scriptContent)))
                    $baseRequirementRule.Add('RunAs32Bit', $requirementRule.runAs32Bit)
                    $baseRequirementRule.Add('EnforceSignatureCheck', $requirementRule.enforceSignatureCheck)
                    $baseRequirementRule.Add('RunAsAccountType', $requirementRule.runAsAccountType)
                    $baseRequirementRule.Add('PowerShellScriptDetectionType', $requirementRule.detectionType)
                }
            }
            $complexRequirementRules += $baseRequirementRule
        }

        if ($null -ne $getValue.AdditionalProperties.InstallExperience)
        {
            $complexInstallExperience = @{}
            $complexInstallExperience.Add('DeviceRestartBehavior', $getValue.AdditionalProperties.installExperience.deviceRestartBehavior)
            $complexInstallExperience.Add('MaxRunTimeInMinutes', $getValue.AdditionalProperties.installExperience.maxRunTimeInMinutes)
            $complexInstallExperience.Add('RunAsAccountType', $getValue.AdditionalProperties.installExperience.runAsAccountType)
        }

        $complexReturnCodes = @()
        foreach ($returnCode in $getValue.AdditionalProperties.returnCodes)
        {
            $complexReturnCodes += @{
                ReturnCode = $returnCode.returnCode
                Type = $returnCode.type
            }
        }

        if ($null -ne $getValue.AdditionalProperties.msiInformation)
        {
            $complexMsiInformation = @{}
            $complexMsiInformation.Add('ProductCode', $getValue.AdditionalProperties.msiInformation.productCode)
            $complexMsiInformation.Add('ProductVersion', $getValue.AdditionalProperties.msiInformation.productVersion)
            $complexMsiInformation.Add('UpgradeCode', $getValue.AdditionalProperties.msiInformation.upgradeCode)
            $complexMsiInformation.Add('RequiresReboot', $getValue.AdditionalProperties.msiInformation.requiresReboot)
            $complexMsiInformation.Add('PackageType', $getValue.AdditionalProperties.msiInformation.packageType)
            $complexMsiInformation.Add('ProductName', $getValue.AdditionalProperties.msiInformation.productName)
            $complexMsiInformation.Add('Publisher', $getValue.AdditionalProperties.msiInformation.publisher)
        }
        #endregion

        $results = @{
            #region resource generator code
            AllowedArchitectures            = $getValue.AdditionalProperties.allowedArchitectures -split ","
            Categories                      = $complexCategories
            Description                     = $getValue.Description
            Developer                       = $getValue.Developer
            DisplayName                     = $getValue.DisplayName
            InformationUrl                  = $getValue.InformationUrl
            InstallCommandLine              = $getValue.AdditionalProperties.installCommandLine
            UninstallCommandLine            = $getValue.AdditionalProperties.uninstallCommandLine
            MinimumSupportedOperatingSystem = $complexMinimumSupportedOperatingSystem
            MinimumFreeDiskSpaceInMB        = $getValue.AdditionalProperties.minimumFreeDiskSpaceInMB
            MinimumMemoryInMB               = $getValue.AdditionalProperties.minimumMemoryInMB
            MinimumNumberOfProcessors       = $getValue.AdditionalProperties.minimumNumberOfProcessors
            MinimumCpuSpeedInMHz            = $getValue.AdditionalProperties.minimumCpuSpeedInMHz
            DetectionRules                  = $complexDetectionRules
            RequirementRules                = $complexRequirementRules
            InstallExperience               = $complexInstallExperience
            ReturnCodes                     = $complexReturnCodes
            MsiInformation                  = $complexMsiInformation
            SetupFilePath                   = $getValue.AdditionalProperties.setupFilePath
            MinimumSupportedWindowsRelease  = $getValue.AdditionalProperties.minimumSupportedWindowsRelease
            DisplayVersion                  = $getValue.AdditionalProperties.displayVersion
            AllowAvailableUninstall         = $getValue.AdditionalProperties.allowAvailableUninstall
            IsFeatured                      = $getValue.IsFeatured
            LargeIcon                       = $complexLargeIcon
            Notes                           = $getValue.Notes
            Owner                           = $getValue.Owner
            PrivacyInformationUrl           = $getValue.PrivacyInformationUrl
            Publisher                       = $getValue.Publisher
            RoleScopeTagIds                 = $getValue.RoleScopeTagIds
            Id                              = $getValue.Id
            Ensure                          = 'Present'
            Credential                      = $Credential
            ApplicationId                   = $ApplicationId
            TenantId                        = $TenantId
            ApplicationSecret               = $ApplicationSecret
            CertificateThumbprint           = $CertificateThumbprint
            ManagedIdentity                 = $ManagedIdentity.IsPresent
            #endregion
        }
        $assignmentsValues = Get-MgBetaDeviceAppManagementMobileAppAssignment -MobileAppId $Id
        $assignmentResult = @()
        if ($assignmentsValues.Count -gt 0)
        {
            $assignmentResult += ConvertFrom-IntunePolicyAssignment -Assignments $assignmentsValues -IncludeDeviceFilter $true
        }
        $results.Add('Assignments', $assignmentResult)

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
        $Id,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String]
        $Developer,

        [Parameter()]
        [System.String]
        $InformationUrl,

        [Parameter()]
        [System.Boolean]
        $IsFeatured,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $LargeIcon,

        [Parameter()]
        [System.String]
        $Notes,

        [Parameter()]
        [System.String]
        $Owner,

        [Parameter()]
        [System.String]
        $PrivacyInformationUrl,

        [Parameter()]
        [System.String]
        $Publisher,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Categories,

        [Parameter()]
        [System.String]
        $InstallCommandLine,

        [Parameter()]
        [System.String]
        $UninstallCommandLine,

        [Parameter()]
        [System.String[]]
        $AllowedArchitectures,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $MinimumSupportedOperatingSystem,

        [Parameter()]
        [System.Int32]
        $MinimumFreeDiskSpaceInMB,

        [Parameter()]
        [System.Int32]
        $MinimumMemoryInMB,

        [Parameter()]
        [System.Int32]
        $MinimumNumberOfProcessors,

        [Parameter()]
        [System.Int32]
        $MinimumCpuSpeedInMHz,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $DetectionRules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $RequirementRules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Rules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $InstallExperience,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $ReturnCodes,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $MsiInformation,

        [Parameter()]
        [System.String]
        $SetupFilePath,

        [Parameter()]
        [System.String]
        $MinimumSupportedWindowsRelease,

        [Parameter()]
        [System.String]
        $DisplayVersion,

        [Parameter()]
        [System.Boolean]
        $AllowAvailableUninstall,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,
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

    Write-Verbose -Message "Setting configuration of the Intune Mobile Apps Win32 App for Windows10 with Id {$Id} and DisplayName {$DisplayName}"

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

    if ($boundParameters.ContainsKey('AllowedArchitectures'))
    {
        if ($boundParameters.AllowedArchitectures -contains 'none' -and $boundParameters.AllowedArchitectures.Count -gt 1)
        {
            throw "AllowedArchitectures cannot contain 'none' when other architectures are specified."
        }

        $boundParameters.Remove('AllowedArchitectures') | Out-Null
        $boundParameters.Add('AllowedArchitectures', ($PSBoundParameters.AllowedArchitectures -join ','))
    }

    if ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Absent')
    {
        Write-Verbose -Message "Creating an Intune Mobile Apps Win32 App for Windows10 with DisplayName {$DisplayName}"
        $boundParameters.Remove("Assignments") | Out-Null

        $createParameters = ([Hashtable]$boundParameters).Clone()
        $createParameters = Rename-M365DSCCimInstanceParameter -Properties $createParameters
        $createParameters.Remove('Id') | Out-Null

        $keys = (([Hashtable]$createParameters).Clone()).Keys
        foreach ($key in $keys)
        {
            if ($null -ne $createParameters.$key -and $createParameters.$key.GetType().Name -like '*CimInstance*')
            {
                $createParameters.$key = Convert-M365DSCDRGComplexTypeToHashtable -ComplexObject $createParameters.$key

                $rulesToProcess = @()
                $isDetectionRule = $false
                if ($key -eq 'DetectionRules')
                {
                    $isDetectionRule = $true
                    $rulesToProcess = $createParameters.DetectionRules
                }
                elseif ($key -eq 'RequirementRules')
                {
                    $rulesToProcess = $createParameters.RequirementRules
                }
                foreach ($rule in $rulesToProcess)
                {
                    $odataType = $rule.OdataType
                    $rule.Remove('OdataType') | Out-Null
                    if ($isDetectionRule)
                    {
                        $rule.Add('@odata.type', "#microsoft.graph.win32LobApp$odataType" + 'Detection')
                    }
                    else
                    {
                        $rule.Add('@odata.type', "#microsoft.graph.win32LobApp$odataType" + 'Requirement')
                    }
                    switch ($odataType)
                    {
                        'FileSystem' {
                            $rule.Add('DetectionType', $rule.FileSystemDetectionType)
                            $rule.Remove('FileSystemDetectionType') | Out-Null
                        }
                        'Registry' {
                            $rule.Add('DetectionType', $rule.RegistryDetectionType)
                            $rule.Remove('RegistryDetectionType') | Out-Null
                        }
                        'PowerShellScript' {
                            $rule.Add('ScriptContent', [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($rule.Script)))
                            $rule.Remove('Script') | Out-Null
                        }
                    }
                }
            }
        }
        #region resource generator code
        $createParameters.Add("@odata.type", "#microsoft.graph.win32LobApp")
        $policy = Invoke-MgGraphRequest -Method POST -Uri "/beta/deviceAppManagement/mobileApps" -Body ($createParameters | ConvertTo-Json -Depth 10)

        if ($PSBoundParameters.ContainsKey('Categories'))
        {
            Update-DeviceAppManagementAppCategory -App $policy -Categories $Categories
        }

        if ($policy.Id)
        {
            $assignmentsHash = ConvertTo-IntuneMobileAppAssignment -IncludeDeviceFilter:$true -Assignments $Assignments
            Update-DeviceAppManagementPolicyAssignment `
                -AppManagementPolicyId $policy.Id `
                -Assignments $assignmentsHash
        }
        #endregion
    }
    elseif ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Updating the Intune Mobile Apps Win32 App for Windows10 with Id {$($currentInstance.Id)}"
        $boundParameters.Remove("Assignments") | Out-Null

        $updateParameters = ([Hashtable]$boundParameters).Clone()
        $updateParameters = Rename-M365DSCCimInstanceParameter -Properties $updateParameters

        $updateParameters.Remove('Id') | Out-Null

        $keys = (([Hashtable]$updateParameters).Clone()).Keys
        foreach ($key in $keys)
        {
            if ($null -ne $createParameters.$key -and $createParameters.$key.GetType().Name -like '*CimInstance*')
            {
                $createParameters.$key = Convert-M365DSCDRGComplexTypeToHashtable -ComplexObject $createParameters.$key

                $rulesToProcess = @()
                $isDetectionRule = $false
                if ($key -eq 'DetectionRules')
                {
                    $isDetectionRule = $true
                    $rulesToProcess = $createParameters.DetectionRules
                }
                elseif ($key -eq 'RequirementRules')
                {
                    $rulesToProcess = $createParameters.RequirementRules
                }
                foreach ($rule in $rulesToProcess)
                {
                    $odataType = $rule.OdataType
                    $rule.Remove('OdataType') | Out-Null
                    if ($isDetectionRule)
                    {
                        $rule.Add('@odata.type', "#microsoft.graph.win32LobApp$odataType" + 'Detection')
                    }
                    else
                    {
                        $rule.Add('@odata.type', "#microsoft.graph.win32LobApp$odataType" + 'Requirement')
                    }
                    switch ($odataType)
                    {
                        'FileSystem' {
                            $rule.Add('DetectionType', $rule.FileSystemDetectionType)
                            $rule.Remove('FileSystemDetectionType') | Out-Null
                        }
                        'Registry' {
                            $rule.Add('DetectionType', $rule.RegistryDetectionType)
                            $rule.Remove('RegistryDetectionType') | Out-Null
                        }
                        'PowerShellScript' {
                            $rule.Add('ScriptContent', [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($rule.Script)))
                            $rule.Remove('Script') | Out-Null
                        }
                    }
                }
            }
        }

        #region resource generator code
        $updateParameters.Add("@odata.type", "#microsoft.graph.win32LobApp")
        Invoke-MgGraphRequest -Method PATCH -Uri "/beta/deviceAppManagement/mobileApps/$($currentInstance.Id)" -Body ($updateParameters | ConvertTo-Json -Depth 10)

        if ($PSBoundParameters.ContainsKey('Categories'))
        {
            Update-DeviceAppManagementAppCategory -App $currentInstance -Categories $Categories -Compare
        }

        $assignmentsHash = ConvertTo-IntuneMobileAppAssignment -IncludeDeviceFilter:$true -Assignments $Assignments
        Update-DeviceAppManagementPolicyAssignment `
            -AppManagementPolicyId $currentInstance.Id `
            -Assignments $assignmentsHash
        #endregion
    }
    elseif ($Ensure -eq 'Absent' -and $currentInstance.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Removing the Intune Mobile Apps Win32 App for Windows10 with Id {$($currentInstance.Id)}"
        #region resource generator code
        Remove-MgBetaDeviceAppManagementMobileApp -MobileAppId $currentInstance.Id
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
        $Id,

        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [System.String]
        $Description,

        [Parameter()]
        [System.String]
        $Developer,

        [Parameter()]
        [System.String]
        $InformationUrl,

        [Parameter()]
        [System.Boolean]
        $IsFeatured,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $LargeIcon,

        [Parameter()]
        [System.String]
        $Notes,

        [Parameter()]
        [System.String]
        $Owner,

        [Parameter()]
        [System.String]
        $PrivacyInformationUrl,

        [Parameter()]
        [System.String]
        $Publisher,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Categories,

        [Parameter()]
        [System.String]
        $InstallCommandLine,

        [Parameter()]
        [System.String]
        $UninstallCommandLine,

        [Parameter()]
        [System.String[]]
        $AllowedArchitectures,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $MinimumSupportedOperatingSystem,

        [Parameter()]
        [System.Int32]
        $MinimumFreeDiskSpaceInMB,

        [Parameter()]
        [System.Int32]
        $MinimumMemoryInMB,

        [Parameter()]
        [System.Int32]
        $MinimumNumberOfProcessors,

        [Parameter()]
        [System.Int32]
        $MinimumCpuSpeedInMHz,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $DetectionRules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $RequirementRules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Rules,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $InstallExperience,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $ReturnCodes,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance]
        $MsiInformation,

        [Parameter()]
        [System.String]
        $SetupFilePath,

        [Parameter()]
        [System.String]
        $MinimumSupportedWindowsRelease,

        [Parameter()]
        [System.String]
        $DisplayVersion,

        [Parameter()]
        [System.Boolean]
        $AllowAvailableUninstall,

        [Parameter()]
        [System.String[]]
        $RoleScopeTagIds,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $Assignments,
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

    Write-Verbose -Message "Testing configuration of the Intune Mobile Apps Win32 App for Windows10 with Id {$Id} and DisplayName {$DisplayName}"

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
        $baseFilter = "isof('microsoft.graph.win32LobApp')"
        if (-not [String]::IsNullOrEmpty($Filter))
        {
            $Filter = "($Filter) and ($baseFilter)"
        }
        else
        {
            $Filter = $baseFilter
        }
        [array]$getValue = Get-MgBetaDeviceAppManagementMobileApp `
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
            if ($null -ne $Results.Categories)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.Categories `
                    -CIMInstanceName 'DeviceManagementMobileAppCategory'

                if (-not [System.String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.Categories = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('Categories') | Out-Null
                }
            }
            if ($null -ne $Results.DetectionRules)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.DetectionRules `
                    -CIMInstanceName 'MicrosoftGraphWin32LobAppDetection'
                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.DetectionRules = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('DetectionRules') | Out-Null
                }
            }
            if ($null -ne $Results.RequirementRules)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.RequirementRules `
                    -CIMInstanceName 'MicrosoftGraphWin32LobAppRequirement'
                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.RequirementRules = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('RequirementRules') | Out-Null
                }
            }
            if ($null -ne $Results.InstallExperience)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.InstallExperience `
                    -CIMInstanceName 'MicrosoftGraphWin32LobAppInstallExperience'
                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.InstallExperience = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('InstallExperience') | Out-Null
                }
            }
            if ($null -ne $Results.ReturnCodes)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.ReturnCodes `
                    -CIMInstanceName 'MicrosoftGraphWin32LobAppReturnCode'
                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.ReturnCodes = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('ReturnCodes') | Out-Null
                }
            }
            if ($null -ne $Results.MsiInformation)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.MsiInformation `
                    -CIMInstanceName 'MicrosoftGraphWin32LobAppMsiInformation'
                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.MsiInformation = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('MsiInformation') | Out-Null
                }
            }
            if ($null -ne $Results.LargeIcon)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.LargeIcon `
                    -CIMInstanceName 'DeviceManagementMimeContent'
                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.LargeIcon = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('LargeIcon') | Out-Null
                }
            }
            if ($null -ne $Results.MinimumSupportedOperatingSystem)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString `
                    -ComplexObject $Results.MinimumSupportedOperatingSystem `
                    -CIMInstanceName 'MicrosoftGraphWindowsMinimumOperatingSystem'
                if (-not [String]::IsNullOrWhiteSpace($complexTypeStringResult))
                {
                    $Results.MinimumSupportedOperatingSystem = $complexTypeStringResult
                }
                else
                {
                    $Results.Remove('MinimumSupportedOperatingSystem') | Out-Null
                }
            }

            if ($Results.Assignments)
            {
                $complexTypeStringResult = Get-M365DSCDRGComplexTypeToString -ComplexObject $Results.Assignments -CIMInstanceName DeviceManagementMobileAppAssignment
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
                -NoEscape @('Assignments', 'Categories', 'DetectionRules', 'RequirementRules', 'LargeIcon', 'InstallExperience', 'ReturnCodes', 'MsiInformation', 'MinimumSupportedOperatingSystem')
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

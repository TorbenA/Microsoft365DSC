[CmdletBinding()]
param(
)
$M365DSCTestFolder = Join-Path -Path $PSScriptRoot `
                        -ChildPath '..\..\Unit' `
                        -Resolve
$CmdletModule = (Join-Path -Path $M365DSCTestFolder `
            -ChildPath '\Stubs\Microsoft365.psm1' `
            -Resolve)
$GenericStubPath = (Join-Path -Path $M365DSCTestFolder `
    -ChildPath '\Stubs\Generic.psm1' `
    -Resolve)
Import-Module -Name (Join-Path -Path $M365DSCTestFolder `
        -ChildPath '\UnitTestHelper.psm1' `
        -Resolve)

$Global:DscHelper = New-M365DscUnitTestHelper -StubModule $CmdletModule `
    -DscResource "IntuneMobileAppsWin32AppWindows10" -GenericStubModule $GenericStubPath
Describe -Name $Global:DscHelper.DescribeHeader -Fixture {
    InModuleScope -ModuleName $Global:DscHelper.ModuleName -ScriptBlock {
        Invoke-Command -ScriptBlock $Global:DscHelper.InitializeScript -NoNewScope
        BeforeAll {

            $secpasswd = ConvertTo-SecureString (New-Guid | Out-String) -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential ('tenantadmin@mydomain.com', $secpasswd)

            Mock -ModuleName M365DSCUtil -CommandName Confirm-M365DSCDependencies -MockWith {
            }

            Mock -CommandName Get-MSCloudLoginConnectionProfile -MockWith {
            }

            Mock -CommandName Reset-MSCloudLoginConnectionProfileContext -MockWith {
            }

            Mock -CommandName Get-PSSession -MockWith {
            }

            Mock -CommandName Remove-PSSession -MockWith {
            }

            Mock -CommandName Update-DeviceAppManagementAppCategory -MockWith {
            }

            Mock -CommandName Update-DeviceAppManagementPolicyAssignment -MockWith {
            }

            Mock -CommandName Update-MgBetaDeviceAppManagementMobileApp -MockWith {
            }

            Mock -CommandName Invoke-MgGraphRequest -MockWith {
                return @{
                    AdditionalProperties = @{
                        '@odata.type' = "#microsoft.graph.win32LobApp"
                        allowedArchitectures = "x86,x64"
                        installCommandLine = "IntuneWinAppUtil.exe -s -t 0"
                        uninstallCommandLine = "IntuneWinAppUtil.exe -s -u -t 0"
                        setupFilePath = "IntuneWinAppUtil.exe"
                        minimumSupportedWindowsRelease = @{
                            v8_0 = $False
                            v8_1 = $False
                            v10_0 = $False
                            v10_1607 = $False
                            v10_1703 = $False
                            v10_1709 = $False
                            v10_1803 = $False
                            v10_1809 = $False
                            v10_1903 = $False
                            v10_1909 = $False
                            v10_2004 = $False
                            v10_2H20 = $False
                            v10_21H1 = $False
                        }
                        displayVersion = "1.0.0.0"
                        allowAvailableUninstall = $False
                        detectionRules = @(
                            @{
                                '@odata.type' = "#microsoft.graph.win32LobAppFileSystemDetection"
                                check32BitOn64System = $False
                                detectionType = "version"
                                detectionValue = "1.0.0.0"
                                fileOrFolderName = "test.exe"
                                operator = "equal"
                                path = "C:\temp\Dev"
                            }
                        )
                        requirementRules = @(
                            @{
                                '@odata.type' = "#microsoft.graph.win32LobAppFileSystemRequirement"
                                check32BitOn64System = $False
                                detectionType = "exists"
                                fileOrFolderName = "test.exe"
                                operator = "notConfigured"
                                path = "C:\temp\Dev"
                            }
                        )
                        returnCodes = @(
                            @{
                                returnCode = 0
                                type = "success"
                            }
                        )
                    }
                    Categories = @(
                        @{
                            Id = "FakeStringValue"
                            DisplayName = "FakeStringValue"
                        }
                    )
                    CommittedContentVersion = "FakeStringValue"
                    DependentAppCount = 25
                    Description = "FakeStringValue"
                    Developer = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    FileName = "FakeStringValue"
                    Id = "FakeStringValue"
                    InformationUrl = "FakeStringValue"
                    IsFeatured = $True
                    LargeIcon = @{
                        Type = "FakeStringValue"
                        Value = [System.Convert]::FromBase64String("VGVzdA==")
                    }
                    Notes = "FakeStringValue"
                    Owner = "FakeStringValue"
                    PrivacyInformationUrl = "FakeStringValue"
                    Publisher = "FakeStringValue"
                    PublishingState = "notPublished"
                    RoleScopeTagIds = @("FakeStringValue")
                    SupersededAppCount = 25
                    SupersedingAppCount = 25
                    UploadState = 25
                }
            }

            Mock -CommandName Remove-MgBetaDeviceAppManagementMobileApp -MockWith {
            }

            Mock -CommandName Get-MgBetaDeviceAppManagementMobileApp -MockWith {
                return @{
                    AdditionalProperties = @{
                        '@odata.type' = "#microsoft.graph.win32LobApp"
                        allowedArchitectures = "x86,x64"
                        installCommandLine = "IntuneWinAppUtil.exe -s -t 0"
                        uninstallCommandLine = "IntuneWinAppUtil.exe -s -u -t 0"
                        setupFilePath = "IntuneWinAppUtil.exe"
                        minimumSupportedWindowsRelease = @{
                            v8_0 = $False
                            v8_1 = $False
                            v10_0 = $False
                            v10_1607 = $False
                            v10_1703 = $False
                            v10_1709 = $False
                            v10_1803 = $False
                            v10_1809 = $False
                            v10_1903 = $False
                            v10_1909 = $False
                            v10_2004 = $False
                            v10_2H20 = $False
                            v10_21H1 = $False
                        }
                        displayVersion = "1.0.0.0"
                        allowAvailableUninstall = $False
                        detectionRules = @(
                            @{
                                '@odata.type' = "#microsoft.graph.win32LobAppFileSystemDetection"
                                check32BitOn64System = $False
                                detectionType = "version"
                                detectionValue = "1.0.0.0"
                                fileOrFolderName = "test.exe"
                                operator = "equal"
                                path = "C:\temp\Dev"
                            }
                        )
                        installExperience = @{
                            deviceRestartBehavior = "suppress"
                            maxRunTimeInMinutes = 60
                            runAsAccountType = "system"
                        }
                        msiInformation = @{
                            productCode = "{00000000-0000-0000-0000-000000000000}"
                            productVersion = "1.0.0.0"
                            upgradeCode = "{00000000-0000-0000-0000-000000000000}"
                            requiresReboot = $False
                            packageType = "dualPurpose"
                            productName = "IntuneWinAppUtil"
                            publisher = "FakeStringValue"
                        }
                        requirementRules = @(
                            @{
                                '@odata.type' = "#microsoft.graph.win32LobAppFileSystemRequirement"
                                check32BitOn64System = $False
                                detectionType = "exists"
                                fileOrFolderName = "test.exe"
                                operator = "notConfigured"
                                path = "C:\temp\Dev"
                            }
                        )
                        returnCodes = @(
                            @{
                                returnCode = 0
                                type = "success"
                            }
                        )
                    }
                    Categories = @(
                        @{
                            Id = "FakeStringValue"
                            DisplayName = "FakeStringValue"
                        }
                    )
                    CommittedContentVersion = "FakeStringValue"
                    DependentAppCount = 25
                    Description = "FakeStringValue"
                    Developer = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    FileName = "FakeStringValue"
                    Id = "FakeStringValue"
                    InformationUrl = "FakeStringValue"
                    IsFeatured = $True
                    LargeIcon = @{
                        Type = "FakeStringValue"
                        Value = [System.Convert]::FromBase64String("VGVzdA==")
                    }
                    Notes = "FakeStringValue"
                    Owner = "FakeStringValue"
                    PrivacyInformationUrl = "FakeStringValue"
                    Publisher = "FakeStringValue"
                    PublishingState = "notPublished"
                    RoleScopeTagIds = @("FakeStringValue")
                    SupersededAppCount = 25
                    SupersedingAppCount = 25
                    UploadState = 25
                }
            }

            Mock -CommandName New-M365DSCConnection -MockWith {
                return "Credentials"
            }

            # Mock Write-M365DSCHost to hide output during the tests
            Mock -CommandName Write-M365DSCHost -MockWith {
            }
            $Script:exportedInstance = $null
            $Script:ExportMode = $false

            Mock -CommandName Get-MgBetaDeviceAppManagementMobileAppAssignment -MockWith {
            }

        }

        # Test contexts
        Context -Name "The IntuneMobileAppsWin32AppWindows10 should exist but it DOES NOT" -Fixture {
            BeforeAll {
                $testParams = @{
                    AllowedArchitectures = @("x86", "x64")
                    Categories = [CimInstance[]]@((New-CimInstance -ClassName MSFT_DeviceManagementMobileAppCategory -Property @{
                        Id = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                    } -ClientOnly))
                    InstallExperience = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppInstallExperience -Property @{
                        DeviceRestartBehavior = "suppress"
                        MaxRunTimeInMinutes = 60
                        RunAsAccountType = "system"
                    } -ClientOnly)
                    MsiInformation = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppMsiInformation -Property @{
                        ProductCode = "{00000000-0000-0000-0000-000000000000}"
                        ProductVersion = "1.0.0.0"
                        UpgradeCode = "{00000000-0000-0000-0000-000000000000}"
                        RequiresReboot = $False
                        PackageType = "dualPurpose"
                        ProductName = "IntuneWinAppUtil"
                        Publisher = "FakeStringValue"
                    } -ClientOnly)
                    DetectionRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemDetection -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "version"
                            DetectionValue = "1.0.0.0"
                            FileOrFolderName = "test.exe"
                            OdataType = "FileSystem"
                            Operator = "equal"
                            Path = "C:\temp\Dev"
                        } -ClientOnly)
                    )
                    RequirementRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemRequirement -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "exists"
                            FileOrFolderName = "test.exe"
                            Operator = "notConfigured"
                            Path = "C:\temp\Dev"
                            OdataType = "FileSystem"
                        } -ClientOnly)
                    )
                    ReturnCodes = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppReturnCode -Property @{
                            ReturnCode = 0
                            Type = "success"
                        } -ClientOnly)
                    )
                    InstallCommandLine = "IntuneWinAppUtil.exe -s -t 0"
                    UninstallCommandLine = "IntuneWinAppUtil.exe -s -u -t 0"
                    SetupFilePath = "IntuneWinAppUtil.exe"
                    Description = "FakeStringValue"
                    Developer = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    Id = "FakeStringValue"
                    InformationUrl = "FakeStringValue"
                    IsFeatured = $True
                    LargeIcon = (New-CimInstance -ClassName MSFT_MicrosoftGraphmimeContent -Property @{
                        Type = "FakeStringValue"
                        Value = "VGVzdA==" # Base64 encoded string for "Test"
                    } -ClientOnly)
                    Notes = "FakeStringValue"
                    Owner = "FakeStringValue"
                    PrivacyInformationUrl = "FakeStringValue"
                    Publisher = "FakeStringValue"
                    RoleScopeTagIds = @("FakeStringValue")
                    Ensure = "Present"
                    Credential = $Credential;
                }

                Mock -CommandName Get-MgBetaDeviceAppManagementMobileApp -MockWith {
                    return $null
                }
            }
            It 'Should return Values from the Get method' {
                (Get-TargetResource @testParams).Ensure | Should -Be 'Absent'
            }
            It 'Should return false from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }
            It 'Should Create the group from the Set method' {
                Set-TargetResource @testParams
                Should -Invoke -CommandName Invoke-MgGraphRequest -Exactly 1
            }
        }

        Context -Name "The IntuneMobileAppsWin32AppWindows10 exists but it SHOULD NOT" -Fixture {
            BeforeAll {
                $testParams = @{
                    AllowedArchitectures = @("x86", "x64")
                    Categories = [CimInstance[]]@((New-CimInstance -ClassName MSFT_DeviceManagementMobileAppCategory -Property @{
                        Id = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                    } -ClientOnly))
                    InstallExperience = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppInstallExperience -Property @{
                        DeviceRestartBehavior = "suppress"
                        MaxRunTimeInMinutes = 60
                        RunAsAccountType = "system"
                    } -ClientOnly)
                    MsiInformation = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppMsiInformation -Property @{
                        ProductCode = "{00000000-0000-0000-0000-000000000000}"
                        ProductVersion = "1.0.0.0"
                        UpgradeCode = "{00000000-0000-0000-0000-000000000000}"
                        RequiresReboot = $False
                        PackageType = "dualPurpose"
                        ProductName = "IntuneWinAppUtil"
                        Publisher = "FakeStringValue"
                    } -ClientOnly)
                    DetectionRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemDetection -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "version"
                            DetectionValue = "1.0.0.0"
                            FileOrFolderName = "test.exe"
                            OdataType = "FileSystem"
                            Operator = "equal"
                            Path = "C:\temp\Dev"
                        } -ClientOnly)
                    )
                    RequirementRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemRequirement -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "exists"
                            FileOrFolderName = "test.exe"
                            Operator = "notConfigured"
                            Path = "C:\temp\Dev"
                            OdataType = "FileSystem"
                        } -ClientOnly)
                    )
                    ReturnCodes = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppReturnCode -Property @{
                            ReturnCode = 0
                            Type = "success"
                        } -ClientOnly)
                    )
                    InstallCommandLine = "IntuneWinAppUtil.exe -s -t 0"
                    UninstallCommandLine = "IntuneWinAppUtil.exe -s -u -t 0"
                    SetupFilePath = "IntuneWinAppUtil.exe"
                    Description = "FakeStringValue"
                    Developer = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    Id = "FakeStringValue"
                    InformationUrl = "FakeStringValue"
                    IsFeatured = $True
                    LargeIcon = (New-CimInstance -ClassName MSFT_MicrosoftGraphmimeContent -Property @{
                        Type = "FakeStringValue"
                        Value = "VGVzdA==" # Base64 encoded string for "Test"
                    } -ClientOnly)
                    Notes = "FakeStringValue"
                    Owner = "FakeStringValue"
                    PrivacyInformationUrl = "FakeStringValue"
                    Publisher = "FakeStringValue"
                    RoleScopeTagIds = @("FakeStringValue")
                    Ensure = "Absent"
                    Credential = $Credential;
                }
            }

            It 'Should return Values from the Get method' {
                (Get-TargetResource @testParams).Ensure | Should -Be 'Present'
            }

            It 'Should return false from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }

            It 'Should Remove the group from the Set method' {
                Set-TargetResource @testParams
                Should -Invoke -CommandName Remove-MgBetaDeviceAppManagementMobileApp -Exactly 1
            }
        }

        Context -Name "The IntuneMobileAppsWin32AppWindows10 Exists and Values are already in the desired state" -Fixture {
            BeforeAll {
                $testParams = @{
                    AllowedArchitectures = @("x86", "x64")
                    Categories = [CimInstance[]]@((New-CimInstance -ClassName MSFT_DeviceManagementMobileAppCategory -Property @{
                        Id = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                    } -ClientOnly))
                    InstallExperience = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppInstallExperience -Property @{
                        DeviceRestartBehavior = "suppress"
                        MaxRunTimeInMinutes = 60
                        RunAsAccountType = "system"
                    } -ClientOnly)
                    MsiInformation = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppMsiInformation -Property @{
                        ProductCode = "{00000000-0000-0000-0000-000000000000}"
                        ProductVersion = "1.0.0.0"
                        UpgradeCode = "{00000000-0000-0000-0000-000000000000}"
                        RequiresReboot = $False
                        PackageType = "dualPurpose"
                        ProductName = "IntuneWinAppUtil"
                        Publisher = "FakeStringValue"
                    } -ClientOnly)
                    DetectionRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemDetection -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "version"
                            DetectionValue = "1.0.0.0"
                            FileOrFolderName = "test.exe"
                            OdataType = "FileSystem"
                            Operator = "equal"
                            Path = "C:\temp\Dev"
                        } -ClientOnly)
                    )
                    RequirementRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemRequirement -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "exists"
                            FileOrFolderName = "test.exe"
                            Operator = "notConfigured"
                            Path = "C:\temp\Dev"
                            OdataType = "FileSystem"
                        } -ClientOnly)
                    )
                    ReturnCodes = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppReturnCode -Property @{
                            ReturnCode = 0
                            Type = "success"
                        } -ClientOnly)
                    )
                    InstallCommandLine = "IntuneWinAppUtil.exe -s -t 0"
                    UninstallCommandLine = "IntuneWinAppUtil.exe -s -u -t 0"
                    SetupFilePath = "IntuneWinAppUtil.exe"
                    Description = "FakeStringValue"
                    Developer = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    Id = "FakeStringValue"
                    InformationUrl = "FakeStringValue"
                    IsFeatured = $True
                    LargeIcon = (New-CimInstance -ClassName MSFT_MicrosoftGraphmimeContent -Property @{
                        Type = "FakeStringValue"
                        Value = "VGVzdA==" # Base64 encoded string for "Test"
                    } -ClientOnly)
                    Notes = "FakeStringValue"
                    Owner = "FakeStringValue"
                    PrivacyInformationUrl = "FakeStringValue"
                    Publisher = "FakeStringValue"
                    RoleScopeTagIds = @("FakeStringValue")
                    Ensure = "Present"
                    Credential = $Credential;
                }
            }

            It 'Should return true from the Test method' {
                Test-TargetResource @testParams | Should -Be $true
            }
        }

        Context -Name "The IntuneMobileAppsWin32AppWindows10 exists and values are NOT in the desired state" -Fixture {
            BeforeAll {
                $testParams = @{
                    AllowedArchitectures = @("x86", "x64")
                    Categories = [CimInstance[]]@((New-CimInstance -ClassName MSFT_DeviceManagementMobileAppCategory -Property @{
                        Id = "FakeStringValue"
                        DisplayName = "FakeStringValue"
                    } -ClientOnly))
                    InstallExperience = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppInstallExperience -Property @{
                        DeviceRestartBehavior = "suppress"
                        MaxRunTimeInMinutes = 60
                        RunAsAccountType = "system"
                    } -ClientOnly)
                    MsiInformation = (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppMsiInformation -Property @{
                        ProductCode = "{00000000-0000-0000-0000-000000000000}"
                        ProductVersion = "1.0.0.0"
                        UpgradeCode = "{00000000-0000-0000-0000-000000000000}"
                        RequiresReboot = $False
                        PackageType = "dualPurpose"
                        ProductName = "IntuneWinAppUtil"
                        Publisher = "FakeStringValue"
                    } -ClientOnly)
                    DetectionRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemDetection -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "version"
                            DetectionValue = "1.0.0.1" # Drift
                            FileOrFolderName = "test.exe"
                            OdataType = "FileSystem"
                            Operator = "equal"
                            Path = "C:\temp\Dev"
                        } -ClientOnly)
                    )
                    RequirementRules = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppFileSystemRequirement -Property @{
                            Check32BitOn64System = $False
                            DetectionType = "exists"
                            FileOrFolderName = "test.exe"
                            Operator = "notConfigured"
                            Path = "C:\temp\Dev"
                            OdataType = "FileSystem"
                        } -ClientOnly)
                    )
                    ReturnCodes = [CimInstance[]]@(
                        (New-CimInstance -ClassName MSFT_MicrosoftGraphWin32LobAppReturnCode -Property @{
                            ReturnCode = 0
                            Type = "success"
                        } -ClientOnly)
                    )
                    InstallCommandLine = "IntuneWinAppUtil.exe -s -t 0"
                    UninstallCommandLine = "IntuneWinAppUtil.exe -s -u -t 0"
                    SetupFilePath = "IntuneWinAppUtil.exe"
                    Description = "FakeStringValue"
                    Developer = "FakeStringValue"
                    DisplayName = "FakeStringValue"
                    Id = "FakeStringValue"
                    InformationUrl = "FakeStringValue"
                    IsFeatured = $True
                    LargeIcon = (New-CimInstance -ClassName MSFT_MicrosoftGraphmimeContent -Property @{
                        Type = "FakeStringValue"
                        Value = "VGVzdA==" # Base64 encoded string for "Test"
                    } -ClientOnly)
                    Notes = "FakeStringValue"
                    Owner = "FakeStringValue"
                    PrivacyInformationUrl = "FakeStringValue"
                    Publisher = "FakeStringValue"
                    RoleScopeTagIds = @("FakeStringValue")
                    Ensure = "Present"
                    Credential = $Credential;
                }
            }

            It 'Should return Values from the Get method' {
                (Get-TargetResource @testParams).Ensure | Should -Be 'Present'
            }

            It 'Should return false from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }

            It 'Should call the Set method' {
                Set-TargetResource @testParams
                Should -Invoke -CommandName Invoke-MgGraphRequest -Exactly 1
            }
        }

        Context -Name 'ReverseDSC Tests' -Fixture {
            BeforeAll {
                $Global:CurrentModeIsExport = $true
                $Global:PartialExportFileName = "$(New-Guid).partial.ps1"
                $testParams = @{
                    Credential = $Credential
                }
            }

            It 'Should Reverse Engineer resource from the Export method' {
                $result = Export-TargetResource @testParams
                $result | Should -Not -BeNullOrEmpty
            }
        }
    }
}

Invoke-Command -ScriptBlock $Global:DscHelper.CleanupScript -NoNewScope

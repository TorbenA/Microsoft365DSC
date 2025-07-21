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
        IntuneMobileAppsWin32AppWindows10 "IntuneMobileAppsWin32AppWindows10-Win32 App - IntuneWinAppUtil.exe"
        {
            AllowAvailableUninstall         = $True;
            AllowedArchitectures            = @("x86","x64","arm64");
            ApplicationId                   = $ApplicationId;
            Assignments                     = @();
            Categories                      = @(
                MSFT_DeviceManagementMobileAppCategory{
                    Id = "2185c6bf-1b3d-4daa-a0bc-79cb4fad9c87"
                    DisplayName = "App Category 1"
                }
            );
            CertificateThumbprint           = $CertificateThumbprint;
            Description                     = "IntuneWinAppUtil.exe";
            DetectionRules                  = @(
                MSFT_MicrosoftGraphWin32LobAppDetection{
                    Check32BitOn64System = $False
                    DetectionValue = "1.0.0.0"
                    FileOrFolderName = "asdf.exe"
                    FileSystemDetectionType = "version"
                    Operator = "equal"
                    Path = "C:\temp\Dev"
                    OdataType = "FileSystem"
                }
                MSFT_MicrosoftGraphWin32LobAppDetection{
                    Check32BitOn64System = $False
                    RegistryDetectionType = "integer"
                    DetectionValue = "100"
                    Operator = "greaterThanOrEqual"
                    OdataType = "Registry"
                    KeyPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion"
                }
                MSFT_MicrosoftGraphWin32LobAppDetection{
                    ProductVersionOperator = "greaterThan" # Drift
                    ProductVersion = "2.0.0.0"
                    ProductCode = "{00000000-0000-0000-0000-000000000000}"
                    OdataType = "ProductCode"
                }
            );
            Developer                       = "";
            DisplayName                     = "Win32 App - IntuneWinAppUtil.exe";
            DisplayVersion                  = "1.0.0.0";
            Ensure                          = "Present";
            Id                              = "63271b78-0fa4-46b8-9ac0-d4b777555dde";
            InstallCommandLine              = "IntuneWinAppUtil.exe -s -t 0";
            IsFeatured                      = $False;
            MinimumCpuSpeedInMHz            = 2500;
            MinimumFreeDiskSpaceInMB        = 1;
            MinimumMemoryInMB               = 200;
            MinimumNumberOfProcessors       = 1;
            MinimumSupportedOperatingSystem = MSFT_MicrosoftGraphWindowsMinimumOperatingSystem{
                V8_0 = $False
                V8_1 = $False
                V10_0 = $False
                V10_1607 = $True
                V10_1703 = $False
                V10_1709 = $False
                V10_1803 = $False
                V10_1809 = $False
                V10_1903 = $False
                V10_1909 = $False
                V10_2004 = $False
                V10_2H20 = $False
                V10_21H1 = $False
            };
            MinimumSupportedWindowsRelease  = "1607";
            Notes                           = "";
            Owner                           = "";
            Publisher                       = "Microsoft";
            RequirementRules                = @(
                MSFT_MicrosoftGraphWin32LobAppRequirement{
                    Check32BitOn64System = $False
                    FileOrFolderName = "asdf.exe"
                    FileSystemDetectionType = "exists"
                    Operator = "notConfigured"
                    Path = "C:\temp\Dev"
                    OdataType = "FileSystem"
                }
                MSFT_MicrosoftGraphWin32LobAppRequirement{
                    Check32BitOn64System = $False
                    RegistryDetectionType = "string"
                    DetectionValue = "1.0.0.0"
                    Operator = "equal"
                    OdataType = "Registry"
                    KeyPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion"
                }
                MSFT_MicrosoftGraphWin32LobAppRequirement{
                    DetectionValue = "Hello World"
                    PowerShellScriptDetectionType = "string"
                    EnforceSignatureCheck = $False
                    Operator = "equal"
                    RunAs32Bit = $False
                    OdataType = "PowerShellScript"
                    Script = "Write-Output `"Hello World`""
                }
            );
            ReturnCodes                     = @(
                MSFT_MicrosoftGraphWin32LobAppReturnCode{
                    ReturnCode = 0
                    Type = "success"
                }
                MSFT_MicrosoftGraphWin32LobAppReturnCode{
                    ReturnCode = 1707
                    Type = "success"
                }
                MSFT_MicrosoftGraphWin32LobAppReturnCode{
                    ReturnCode = 3010
                    Type = "softReboot"
                }
                MSFT_MicrosoftGraphWin32LobAppReturnCode{
                    ReturnCode = 1641
                    Type = "hardReboot"
                }
                MSFT_MicrosoftGraphWin32LobAppReturnCode{
                    ReturnCode = 1618
                    Type = "retry"
                }
            );
            RoleScopeTagIds                 = @("0");
            SetupFilePath                   = "IntuneWinAppUtil.exe";
            TenantId                        = $TenantId;
            UninstallCommandLine            = "IntuneWinAppUtil.exe -t -s 0";
        }
    }
}

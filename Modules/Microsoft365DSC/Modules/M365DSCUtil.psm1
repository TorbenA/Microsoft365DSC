#region Session Objects
$Global:SessionSecurityCompliance = $null
[hashtable]$Script:M365DSCTelemetryConnectionToGraphParams = @{}
#endregion

# Automatically initialize accelerator on module import
Initialize-M365DSCDllLoader -ErrorAction SilentlyContinue

$Script:M365DSCWorkloads = @('AAD', 'ADO', 'AZURE', 'COMMERCE', 'DEFENDER', 'EXO', 'FABRIC', 'INTUNE', 'O365', 'OD', 'PLANNER', 'PP', 'SC', 'SENTINEL', 'SH', 'SPO', 'TEAMS')
$Script:IsPowerShellCore = $PSVersionTable.PSEdition -eq 'Core'
$Script:M365DSCStringReplacementMap = @{}

<#
.DESCRIPTION
    This function retrieves the resources available in the M365DSC project based on the specified export mode.

.FUNCTIONALITY
    Public

.PARAMETER Mode
    Specifies the mode of the export. Valid values are 'Default' and 'Full'.
    - 'Default' includes only configuration resources.
    - 'Full' includes all resources, both configuration and data.

.PARAMETER ExcludeConfigurationResources
    If specified, configuration resources will be excluded from the results. Works only for the 'Full' mode.

.EXAMPLE
    Get-M365DSCResourcesByExportMode -Mode 'Default'

    This command retrieves all resources that are available in the Default export mode.

.EXAMPLE
    Get-M365DSCResourcesByExportMode -Mode 'Full'

    This command retrieves all resources that are available in the Full export mode.

.OUTPUTS
    [System.String[]] - An array of resource names that match the specified export mode.
#>
function Get-M365DSCResourcesByExportMode
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('Default', 'Full')]
        [System.String]
        $Mode,

        [Parameter(Mandatory = $false)]
        [switch]
        $ExcludeConfigurationResources
    )

    $resourceSettings = Get-M365DSCResourceSettings
    $resources = [System.Collections.Generic.List[System.String]]::new($resourceSettings.Keys.Count)
    foreach ($resource in $resourceSettings.Keys)
    {
        if ($Mode -eq 'Default' -and $resourceSettings[$resource].mode -eq 'Configuration')
        {
            $resources.Add($resource)
        }
        elseif ($Mode -eq 'Full')
        {
            if ($ExcludeConfigurationResources -and $resourceSettings[$resource].mode -eq 'Configuration')
            {
                continue
            }
            $resources.Add($resource)
        }
    }

    return $resources.ToArray()
}

<#
.Description
This function retrieves a Teams team by its name

.Functionality
Internal
#>
function Get-TeamByName
{
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $TeamName
    )

    try
    {
        $loopCounter = 0
        do
        {
            $team = Get-Team -DisplayName $TeamName | Where-Object -FilterScript { $_.DisplayName -eq [System.Net.WebUtility]::UrlDecode($TeamName) }
            if ($null -eq $team)
            {
                Start-Sleep 5
            }
            $loopCounter += 1
            if ($loopCounter -gt 5)
            {
                break
            }
        } while ($null -eq $team)

        if ($null -eq $team)
        {
            throw "Team with Name $TeamName doesn't exist in tenant"
        }
        elseif ($teams.Length -gt 1)
        {
            Write-Warning -Message "More than one Team with name {$TeamName} was found. This could prevent your configuration from compiling properly."
        }
        return $team
    }
    catch
    {
        return $null
    }
}

<#
.Description
    This function converts a parameter hashtable to a string, for outputting to screen

.Functionality
    Internal
#>
function Convert-M365DscHashtableToString
{
    param
    (
        [Parameter()]
        [System.Collections.Hashtable]
        $Hashtable
    )

    return [Microsoft365DSC.Converter.HashtableConverter]::ToString($Hashtable)
}

<#
.Description
This function checks if the specified cmdlet is available or not

.Functionality
Internal
#>
function Confirm-ImportedCmdletIsAvailable
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $CmdletName
    )

    try
    {
        $CmdletIsAvailable = (Get-Command -Name $CmdletName -ErrorAction SilentlyContinue)
        if ($CmdletIsAvailable)
        {
            return $true
        }
        else
        {
            return $false
        }
    }
    catch
    {
        return $false
    }
}

<#
.Description
This function compares two arrays with PSCustomObject objects

.Functionality
Internal, Hidden
#>
function Compare-PSCustomObjectArrays
{
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Object[]]
        $DesiredValues,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Object[]]
        $CurrentValues
    )
    $DriftedProperties = @()
    foreach ($DesiredEntry in $DesiredValues)
    {
        $Properties = $DesiredEntry.PSObject.Properties
        $KeyProperty = $Properties.Name[0]

        $EquivalentEntryInCurrent = $CurrentValues | Where-Object -FilterScript { $_.$KeyProperty -eq $DesiredEntry.$KeyProperty }
        if ($null -eq $EquivalentEntryInCurrent)
        {
            $result = @{
                Property     = $DesiredEntry
                PropertyName = $KeyProperty
                Desired      = $DesiredEntry.$KeyProperty
                Current      = $null
            }
            $DriftedProperties += $result
        }
        else
        {
            foreach ($property in $Properties)
            {
                $propertyName = $property.Name

                if ((-not [System.String]::IsNullOrEmpty($DesiredEntry.$PropertyName) -and -not [System.String]::IsNullOrEmpty($EquivalentEntryInCurrent.$PropertyName)) -and `
                    $DesiredEntry.$PropertyName -ne $EquivalentEntryInCurrent.$PropertyName)
                {
                    $drift = $true
                    if ($DesiredEntry.$PropertyName.GetType().Name -eq 'String' -and $DesiredEntry.$PropertyName.Contains('$OrganizationName'))
                    {
                        if ($DesiredEntry.$PropertyName.Split('@')[0] -eq $EquivalentEntryInCurrent.$PropertyName.Split('@')[0])
                        {
                            $drift = $false
                        }
                    }
                    if ($drift)
                    {
                        $result = @{
                            Property     = $DesiredEntry
                            PropertyName = $PropertyName
                            Desired      = $DesiredEntry.$PropertyName
                            Current      = $EquivalentEntryInCurrent.$PropertyName
                        }
                        $DriftedProperties += $result
                    }
                }
            }
        }
    }

    foreach ($currentEntry in $currentValues)
    {
        if ($currentEntry.GetType().Name -eq 'PSCustomObject')
        {
            $fixedEntry = @{}
            $currentEntry.psobject.properties | ForEach-Object { $fixedEntry[$_.Name] = $_.Value }
        }
        else
        {
            $fixedEntry = $currentEntry
        }
        $KeyProperty = Get-M365DSCCIMInstanceKey -CIMInstance $fixedEntry

        $EquivalentEntryInDesired = $DesiredValues | Where-Object -FilterScript { $_.$KeyProperty -eq $fixedEntry.$KeyProperty }
        if ($null -eq $EquivalentEntryInDesired)
        {
            $result = @{
                Property     = $fixedEntry
                PropertyName = $KeyProperty
                Desired      = $fixedEntry.$KeyProperty
                Current      = $null
            }
            $DriftedProperties += $result
        }
        else
        {
            foreach ($property in $Properties)
            {
                $propertyName = $property.Name

                if ((-not [System.String]::IsNullOrEmpty($fixedEntry.$PropertyName) -and -not [System.String]::IsNullOrEmpty($EquivalentEntryInDesired.$PropertyName)) -and `
                    $fixedEntry.$PropertyName -ne $EquivalentEntryInDesired.$PropertyName)
                {
                    $drift = $true
                    if ($fixedEntry.$PropertyName.GetType().Name -eq 'String' -and $fixedEntry.$PropertyName.Contains('$OrganizationName'))
                    {
                        if ($fixedEntry.$PropertyName.Split('@')[0] -eq $EquivalentEntryInDesired.$PropertyName.Split('@')[0])
                        {
                            $drift = $false
                        }
                    }
                    if ($drift)
                    {
                        $result = @{
                            Property     = $fixedEntry
                            PropertyName = $PropertyName
                            Desired      = $fixedEntry.$PropertyName
                            Current      = $EquivalentEntryInDesired.$PropertyName
                        }
                        $DriftedProperties += $result
                    }
                }
            }
        }
    }

    return $DriftedProperties
}


<#
.Description
This function retrieves the current tenant's name based on received authentication parameters.

.Functionality
Internal
#>
function Get-M365DSCTenantNameFromParameterSet
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [System.Collections.HashTable]
        $ParameterSet
    )

    if ($ParameterSet.ContainsKey('TenantId'))
    {
        return $ParameterSet.TenantId
    }
    elseif ($ParameterSet.ContainsKey('Credential'))
    {
        try
        {
            $tenantName = $ParameterSet.Credential.Username.Split('@')[1]
            return $tenantName
        }
        catch
        {
            return $null
        }
    }
}

<#
.DESCRIPTION
    This function converts a property value to an array of specified element type.

.FUNCTIONALITY
    Internal
#>
function Get-M365DSCArrayFromProperty
{
    [CmdletBinding()]
    [OutputType([System.Array])]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [System.Object]
        $PropertyValue,

        [Parameter(Mandatory = $false)]
        [System.Type]
        $ElementType = [System.Object]
    )

    $array = [System.Array]::CreateInstance($ElementType, 0)
    foreach ($item in $PropertyValue)
    {
        $array += $item
    }

    ,$array
}

<#
.Description
This function tests if the DSC hashtables have the same values

.Functionality
Public
#>
function Test-M365DSCParameterState
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [HashTable]
        $CurrentValues,

        [Parameter(Mandatory = $true, Position = 2)]
        [Object]
        $DesiredValues,

        [Parameter(Position = 3)]
        [Array]
        $ValuesToCheck,

        [Parameter(Position = 4)]
        [System.String]
        $Source = 'Generic',

        [Parameter(Position = 5)]
        [System.Collections.Hashtable]
        $IncludedDrifts,

        [Parameter(Position = 6)]
        [switch]
        $NoEventMessage,

        [Parameter(Position = 7)]
        [switch]
        $NoDriftReset,

        [Parameter(Position = 8)]
        [System.String[]]
        $ExcludedProperties
    )

    $startTime = [System.DateTime]::Now
    if ($null -eq $Global:AllDrifts -or -not $NoDriftReset)
    {
        $Global:AllDrifts = @{
            DriftInfo     = @()
            CurrentValues = @{}
            DesiredValues = @{}
        }
        $Global:PotentialDrifts = @()
    }

    $returnValue = $true
    $TenantName = Get-M365DSCTenantNameFromParameterSet -ParameterSet $DesiredValues

    #region Telemetry
    if (Test-IsM365DSCTelemetryEnabled)
    {
        $data = [System.Collections.Generic.Dictionary[[System.String], [System.String]]]::new()
        $data.Add('Resource', "$Source")
        $data.Add('Method', 'Test-TargetResource')

        $dataEvaluation = [System.Collections.Generic.Dictionary[[System.String], [System.String]]]::new()
        $dataEvaluation.Add('Resource', "$Source")
        $dataEvaluation.Add('Method', 'Test-TargetResource')
        $dataEvaluation.Add('Tenant', $TenantName)

        $ConnectionMode = Get-M365DSCAuthenticationMode $DesiredValues
        $dataEvaluation.Add('ConnectionMode', $ConnectionMode)
        # Most likely unnecessary - Keep as a comment for now
        # TODO: Measure performance impact
        <#
        for ($i = 0; $i -lt $ValuesToCheck.Length; $i++)
        {
            if ($ValuesToCheck[$i] -eq 'Verbose')
            {
                $ValuesToCheck.RemoveAt($i)
                break
            }
        }
        #>
        $dataEvaluation.Add('Parameters', $ValuesToCheck -join "`r`n")
        $dataEvaluation.Add('ParametersCount', $ValuesToCheck.Length)
        Add-M365DSCTelemetryEvent -Type 'DriftEvaluation' -Data $dataEvaluation
    }

    $compareResult = [Microsoft365DSC.Compare.SimpleObjectComparer]::Compare($CurrentValues, $DesiredValues, $ValuesToCheck, $IncludedDrifts, $NoEventMessage, $NoDriftReset, $ExcludedProperties)
    $driftedParameters = $compareResult.DriftedParameters
    $driftObject = $compareResult.DriftObject
    $returnValue = $compareResult.TestResult

    $includeNonDriftsInformation = $false
    try
    {
        $includeNonDriftsInformation = [System.Environment]::GetEnvironmentVariable('M365DSCEventLogIncludeNonDrifted', `
                [System.EnvironmentVariableTarget]::Machine)
    }
    catch
    {
        Write-Verbose -Message $_
    }

    if ($returnValue -eq $false -or $DriftedParameters.Keys.Length -gt 0)
    {
        $EventMessage = [System.Text.StringBuilder]::New()
        $EventMessage.Append("<M365DSCEvent>`r`n") | Out-Null
        Write-Verbose -Message "Found Tenant Name: $TenantName"

        $LCMState = $null
        try
        {
            if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) `
                -and $null -eq $Script:LCMInfo)
            {
                $Script:LCMInfo = Get-DscLocalConfigurationManager -ErrorAction Stop

                if ($Script:LCMInfo.LCMStateDetail -eq 'LCM is performing a consistency check.' -or `
                        $Script:LCMInfo.LCMStateDetail -eq 'LCM exécute une vérification de cohérence.' -or `
                        $Script:LCMInfo.LCMStateDetail -eq 'LCM führt gerade eine Konsistenzüberprüfung durch.')
                {
                    $LCMState = 'ConsistencyCheck'
                }
                elseif ($Script:LCMInfo.LCMStateDetail -eq 'LCM is testing node against the configuration.')
                {
                    $LCMState = 'ManualTestDSCConfiguration'
                }
                elseif ($Script:LCMInfo.LCMStateDetail -eq 'LCM is applying a new configuration.' -or `
                        $Script:LCMInfo.LCMStateDetail -eq 'LCM applique une nouvelle configuration.')
                {
                    $LCMState = 'Initial'
                }
            }
            else
            {
                $LCMState = 'Unauthorized'
            }
        }
        catch
        {
            Write-Verbose -Message $_.Exception
        }
        $EventMessage.Append("    <ConfigurationDrift Source=`"$Source`" TenantId=`"$TenantName`"") | Out-Null
        if (-not [System.String]::IsNullOrEmpty($LCMState))
        {
            $EventMessage.Append(" LCMState=`"" + $LCMState + "`"") | Out-Null
        }
        $EventMessage.Append(">`r`n") | Out-Null
        $EventMessage.Append("        <ParametersNotInDesiredState>`r`n") | Out-Null

        $driftedData = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
        $driftedData.Add('Tenant', $TenantName)
        $DriftObject.Add('Tenant', $TenantName)
        $driftedData.Add('Resource', $source.Split('_')[1])
        $DriftObject.Add('Resource', $source.Split('_')[1])

        # If custom App Insights is specified, allow for the current and desired values to be captured;
        # ISSUE #1222
        if ($null -ne $env:M365DSCTelemetryInstrumentationKey -and `
                $env:M365DSCTelemetryInstrumentationKey -ne 'bc5aa204-0b1e-4499-a955-d6a639bdb4fa' -and `
                $env:M365DSCTelemetryInstrumentationKey -ne 'e670af5d-fd30-4407-a796-8ad30491ea7a')
        {
            $driftedData.Add('CurrentValues', $CurrentValues)
            $driftedData.Add('DesiredValues', $DesiredValues)
        }
        #endregion
        $telemetryDriftedParameters = ''
        foreach ($key in $DriftedParameters.Keys)
        {
            Write-Verbose -Message "Detected Drifted Parameter [$Source]$key"
            $telemetryDriftedParameters += $key + "`r`n"
            $EventMessage.Append("            <Param Name=`"$key`">" + $DriftedParameters.$key + "</Param>`r`n") | Out-Null
        }

        $driftedData.Add('Parameters', $telemetryDriftedParameters)
        Add-M365DSCTelemetryEvent -Type 'DriftInfo' -Data $driftedData
        $EventMessage.Append("        </ParametersNotInDesiredState>`r`n") | Out-Null
        $EventMessage.Append("    </ConfigurationDrift>`r`n") | Out-Null
        $EventMessage.Append("    <DesiredValues>`r`n") | Out-Null
        foreach ($Key in $DesiredValues.Keys)
        {
            $Value = $DesiredValues.$Key
            if ([System.String]::IsNullOrEmpty($Value))
            {
                $Value = "`$null"
            }
            $EventMessage.Append("        <Param Name =`"$key`">$Value</Param>`r`n") | Out-Null
            $DriftObject.DesiredValues.Add($key, $value)
        }
        $EventMessage.Append("    </DesiredValues>`r`n") | Out-Null
        $EventMessage.Append("    <CurrentValues>`r`n") | Out-Null
        foreach ($Key in $CurrentValues.Keys)
        {
            $Value = $CurrentValues.$Key
            if ([System.String]::IsNullOrEmpty($Value))
            {
                $Value = "`$null"
            }
            $EventMessage.Append("        <Param Name =`"$key`">$Value</Param>`r`n") | Out-Null
            $DriftObject.CurrentValues.Add($key, $value)
        }
        $EventMessage.Append("    </CurrentValues>`r`n") | Out-Null
        $EventMessage.Append('</M365DSCEvent>') | Out-Null
        foreach ($drift in $DriftObject.DriftInfo)
        {
            $Global:AllDrifts.DriftInfo += @{
                PropertyName = $drift.PropertyName
                CurrentValue = $drift.CurrentValue
                DesiredValue = $drift.DesiredValue
            }
        }
        if (-not $NoEventMessage)
        {
            Add-M365DSCEvent -Message $EventMessage.ToString() -EventType 'Drift' -EntryType 'Warning' `
                -EventID 1 -Source $Source
        }
        $Global:CCMCurrentDriftInfo = $DriftObject
    }
    elseif ($includeNonDriftsInformation -eq $true)
    {
        # Include details about non-drifted resources.
        $EventMessage = [System.Text.StringBuilder]::New()
        $EventMessage.Append("<M365DSCEvent>`r`n") | Out-Null
        $EventMessage.Append("    <ConfigurationDrift Source=`"$Source`" />`r`n") | Out-Null
        $EventMessage.Append("    <DesiredValues>`r`n") | Out-Null
        foreach ($Key in $DesiredValues.Keys)
        {
            $Value = $DesiredValues.$Key
            if ([System.String]::IsNullOrEmpty($Value))
            {
                $Value = "`$null"
            }
            $EventMessage.Append("        <Param Name =`"$key`">$Value</Param>`r`n") | Out-Null
        }
        $EventMessage.Append("    </DesiredValues>`r`n") | Out-Null
        $EventMessage.Append('</M365DSCEvent>') | Out-Null
        Add-M365DSCEvent -Message $EventMessage.ToString() -EventType 'NonDrift' -EntryType 'Information' `
            -EventID 2 -Source $Source
    }

    $timeTaken = [System.DateTime]::Now.Subtract($startTime).TotalMilliseconds
    $data  = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add('Resource', $Source)
    $data.Add('Method', 'Test-M365DSCParameterState')
    $data.Add('TimeTakenMilliseconds', $timeTaken)
    $data.Add('Tenant', $TenantName)
    $data.Add('ParametersCount', $KeyList.Count)

    Add-M365DSCTelemetryEvent -Type 'ResourceTesting' `
        -Data $data
    return $returnValue
}

<#
.Description
    Centralized method to evaluate the result of the various Test-TargetResource functions

.PARAMETER PostProcessing
    Optional Func delegate that allows custom processing of the DesiredValues, CurrentValues and ValuesToCheck.
    The function receives three hashtable parameters: DesiredValues, CurrentValues (from Get-TargetResource) and ValuesToCheck.
    Additionally, it gets an array of objects as PostProcessingArgs.
    The delegate must return a Tuple[Hashtable, Hashtable, Hashtable] where Item1 is the processed DesiredValues, Item2 is the processed CurrentValues and Item3 is the processed ValuesToCheck.

.FUNCTIONALITY
    Internal
#>
function Test-M365DSCTargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    [OutputType([System.Collections.Hashtable], ParameterSetName = 'PassThru')]
    param(
        [Parameter()]
        $DesiredValues,

        [Parameter()]
        [System.String]
        $ResourceName,

        [Parameter()]
        [System.String[]]
        $ExcludedProperties,

        [Parameter()]
        [System.String[]]
        $IncludedProperties,

        [Parameter()]
        [System.Func[Hashtable, Hashtable, Hashtable, [Object[]], Tuple[Hashtable, Hashtable, Hashtable]]]
        $PostProcessing,

        [Parameter()]
        [System.Object[]]
        $PostProcessingArgs = @(),

        [Parameter(
            ParameterSetName = 'PassThru'
        )]
        [switch]
        $PassThru
    )

    $Global:AllDrifts = @{
        DriftInfo     = @()
        CurrentValues = @{}
        DesiredValues = @{}
    }
    $Global:PotentialDrifts = @()

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    if ($null -eq (Get-Module -Name 'M365DSCCompare'))
    {
        $compareModulePath = Join-Path -Path $PSScriptRoot -ChildPath "M365DSCCompare.psm1"
        Import-Module -Name $compareModulePath -Force
    }

    # Retrieve the primary keys of the given resource and remove them from the list of values to check.
    $currentPath = $PSScriptRoot
    if ($null -eq $Script:M365DSCSchema)
    {
        $schemaPath = Join-Path -Path $currentPath -ChildPath '..\SchemaDefinition.json'
        $schemaJSON = Get-Content $schemaPath -Raw
        $Script:M365DSCSchema = ConvertFrom-Json $schemaJSON
    }
    $resourceDefinition = $Script:M365DSCSchema | Where-Object -FilterScript { $_.ClassName -eq "MSFT_$ResourceName" }
    $resourceKeys = $resourceDefinition.Parameters | Where-Object -FilterScript { $_.Option -eq 'Key' }

    $keyStrings = @()
    foreach ($resourceKey in $resourceKeys)
    {
        $keyName = $resourceKey.Name
        $keyStrings += "$keyName {$($DesiredValues.$keyName)}"
    }
    $finalString = $keyStrings -join ' and '

    $Verbose = $false
    if ($DesiredValues.Verbose -eq $true)
    {
        $Verbose = $true
    }

    Write-Verbose -Message "Testing configuration of the $ResourceName with $finalString" -Verbose:$Verbose

    $CurrentValues = & MSFT_$ResourceName\Get-TargetResource @DesiredValues

    $testTargetResource = Compare-M365DSCResourceState -ResourceName $ResourceName `
        -DesiredValues $DesiredValues `
        -CurrentValues $CurrentValues `
        -ExcludedProperties $ExcludedProperties `
        -IncludedProperties $IncludedProperties `
        -PostProcessing $PostProcessing `
        -PostProcessingArgs $PostProcessingArgs

    if (-not $testTargetResource)
    {
        $TenantName = Get-M365DSCTenantNameFromParameterSet -ParameterSet $DesiredValues
        Write-M365DSCDriftsToEventLog -Drifts $Global:AllDrifts `
                                      -ResourceName $ResourceName `
                                      -TenantName $TenantName `
                                      -CurrentValues $CurrentValues `
                                      -DesiredValues $DesiredValues `
                                      -Verbose:$Verbose
    }

    Write-Verbose -Message "Test-M365DSCTargetResource returned $testTargetResource" -Verbose:$Verbose

    if ($PassThru)
    {
        return @{
            ResourceName = $ResourceName
            CurrentValues = $CurrentValues
            DesiredValues = $DesiredValues
            TestTargetResource = $testTargetResource
        }
    }

    return $testTargetResource
}

<#
.DESCRIPTION
    Sets the script scoped variable that holds all the M365DSC resources.

.PARAMETER DscResourceDictionary
    A dictionary containing all the M365DSC resources.

.FUNCTIONALITY
    Internal
#>
function Set-M365DSCAllResourcesDictionary
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $DscResourceDictionary
    )

    $Script:AllM365DSCResources = $DscResourceDictionary
}

<#
.DESCRIPTION
    Retrieves the script scoped variable that holds all the M365DSC resources.

.FUNCTIONALITY
    Internal
#>
function Get-M365DSCAllResourcesDictionary
{
    [CmdletBinding()]
    param()

    $Script:AllM365DSCResources
}

<#
.DESCRIPTION
    Initializes the script scoped variable that holds all the M365DSC resources.

.FUNCTIONALITY
    Internal
#>
function Initialize-M365DSCAllResourcesDictionary
{
    [CmdletBinding()]
    param()

    if ($null -eq $Script:AllM365DSCResources -and -not $Global:IsTestEnvironment)
    {
        $Script:AllM365DSCResources = [System.Collections.Generic.Dictionary[System.String, System.Object]]::new([System.StringComparer]::InvariantCultureIgnoreCase)
        $resources = Get-DscResourceV2 -Module 'Microsoft365DSC'
        foreach ($resource in $resources)
        {
            $Script:AllM365DSCResources.Add($resource.Name, $resource)
        }
    }
}

<#
.Description
This is the main Microsoft365DSC.Reverse function that extracts the DSC configuration from an existing Microsoft 365 Tenant.

.Parameter LaunchWebUI
Adding this parameter will open the WebUI in a browser.

.Parameter Path
Specifies the path in which the exported DSC configuration should be stored.

.Parameter FileName
Specifies the name of the file in which the exported DSC configuration should be stored.

.Parameter ConfigurationName
Specifies the name of the configuration that will be generated.

.Parameter Components
Specifies the components for which an export should be created.

.Parameter ExcludeComponents
Specifies the components to skip when creating the export

.Parameter Workloads
Specifies the workload for which an export should be created for all resources.

.Parameter Mode
Specifies the mode of the export: Default or Full.

.Parameter GenerateInfo
Specifies if each exported resource should get a link to the Wiki article of the resource.

.Parameter ApplicationId
Specifies the application id to be used for authentication.

.Parameter ApplicationSecret
Specifies the application secret of the application to be used for authentication.

.Parameter TenantId
Specifies the id of the tenant.

.Parameter CertificateThumbprint
Specifies the thumbprint to be used for authentication.

.Parameter Credential
Specifies the credentials to be used for authentication.

.Parameter CertificatePassword
Specifies the password of the PFX file which is used for authentication.

.Parameter CertificatePath
Specifies the path of the PFX file which is used for authentication.

.Parameter Filters
Specifies resource level filters to apply in order to reduce the number of instances exported.

.PARAMETER AccessTokens
    Specifies the access token to use for authentication.

.Parameter ManagedIdentity
Specifies use of managed identity for authentication.

.Parameter Validate
Specifies that the configuration needs to be validated for conflicts or issues after its extraction is completed.

.PARAMETER Parallel
    Specifies that the export is executed in parallel.

.PARAMETER TokenReplacement
    Specifies the hashtable to use for token replacement. Key is the value to replace, and the value is the variable to use for replacement without the '$' sign.

.PARAMETER WithStatistics
    Specifies that statistics about the export should be shown after completion.

.Example
Export-M365DSCConfiguration -Components @("AADApplication", "AADConditionalAccessPolicy", "AADGroupsSettings") -Credential $Credential

.Example
Export-M365DSCConfiguration -Mode 'Default' -ApplicationId '2560bb7c-bc85-415f-a799-841e10ec4f9a' -TenantId 'contoso.sharepoint.com' -ApplicationSecret 'abcdefghijkl'

.Example
Export-M365DSCConfiguration -Components @("AADApplication", "AADConditionalAccessPolicy", "AADGroupsSettings") -Credential $Credential -Path 'C:\DSC' -FileName 'MyConfig.ps1'

.Example
Export-M365DSCConfiguration -Credential $Credential -Filters @{AADApplication = "DisplayName eq 'MyApp'"} -TokenReplacement @{ 'alternate-email.onmicrosoft.com' = 'AlternateEmail' }

.Example
Export-M365DSCConfiguration -Workloads @("SPO") -ExcludeComponents @("SPOPropertyBag") -Credential $Credential

.Functionality
Public
#>
function Export-M365DSCConfiguration
{
    [CmdletBinding(DefaultParameterSetName = 'Export')]
    param
    (
        [Parameter(ParameterSetName = 'WebUI')]
        [Switch]
        $LaunchWebUI,

        [Parameter(ParameterSetName = 'Export')]
        [System.String]
        $Path,

        [Parameter(ParameterSetName = 'Export')]
        [System.String]
        $FileName,

        [Parameter(ParameterSetName = 'Export')]
        [System.String]
        $ConfigurationName,

        [Parameter(ParameterSetName = 'Export')]
        [System.String[]]
        $Components,

        [Parameter(ParameterSetName = 'Export')]
        [System.String[]]
        $ExcludeComponents,

        [Parameter(ParameterSetName = 'Export')]
        [ValidateSet('AAD', 'ADO', 'AZURE', 'COMMERCE', 'DEFENDER', 'EXO', 'FABRIC', 'INTUNE', 'O365', 'OD', 'PLANNER', 'PP', 'SC', 'SENTINEL', 'SH', 'SPO', 'TEAMS')]
        [System.String[]]
        $Workloads,

        [Parameter(ParameterSetName = 'Export')]
        [ValidateSet('Default', 'Full')]
        [System.String]
        $Mode = 'Default',

        [Parameter(ParameterSetName = 'Export')]
        [System.Boolean]
        $GenerateInfo = $false,

        [Parameter(ParameterSetName = 'Export')]
        [System.Collections.Hashtable]
        $Filters,

        [Parameter(ParameterSetName = 'Export')]
        [System.String]
        $ApplicationId,

        [Parameter(ParameterSetName = 'Export')]
        [ValidateScript({
            $invalid = $false
            $parseGuid = [System.Guid]::Empty
            if ([System.Guid]::TryParse($_, [ref]$parseGuid))
            {
                throw "Please provide the tenant name (e.g., contoso.onmicrosoft.com) for TenantId instead of its GUID."
            }
            $invalid = $_ -notmatch ".onmicrosoft."
            if ($invalid)
            {
                Write-Warning -Message "We recommend providing the TenantId property in the format of <tenant>.onmicrosoft.*"
            }
            return $true
        })]
        [System.String]
        $TenantId,

        [Parameter(ParameterSetName = 'Export')]
        [System.String]
        $ApplicationSecret,

        [Parameter(ParameterSetName = 'Export')]
        [System.String]
        $CertificateThumbprint,

        [Parameter(ParameterSetName = 'Export')]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter(ParameterSetName = 'Export')]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter(ParameterSetName = 'Export')]
        [System.String]
        $CertificatePath,

        [Parameter(ParameterSetName = 'Export')]
        [Switch]
        $ManagedIdentity,

        [Parameter(ParameterSetName = 'Export')]
        [System.String[]]
        $AccessTokens,

        [Parameter(ParameterSetName = 'Export')]
        [Switch]
        $Validate,

        [Parameter(ParameterSetName = 'Export')]
        [Switch]
        $Parallel,

        [Parameter(ParameterSetName = 'Export')]
        [System.Collections.Hashtable]
        $TokenReplacement,

        [Parameter(ParameterSetName = 'Export')]
        [Switch]
        $WithStatistics
    )

    $currentStartDateTime = [System.DateTime]::Now
    $Global:M365DSCExportInProgress = $true
    $Global:MaximumFunctionCount = 32767

    # Define the exported resource instances' names Global variable
    $Global:M365DSCExportedResourceInstancesNames = @()

    # LaunchWebUI specified, launching that now
    if ($LaunchWebUI)
    {
        Write-Output -InputObject "Launching web page 'https://export.microsoft365dsc.com'"
        explorer 'https://export.microsoft365dsc.com'
        return
    }

    # Suppress Progress overlays
    $Global:ProgressPreference = 'SilentlyContinue'

    # Check ErrorActionPreference - Azure DevOps and other Pipeline environments set it to 'Stop' by default
    if ($ErrorActionPreference -eq 'Stop' -and -not $PSBoundParameters.ContainsKey('ErrorAction'))
    {
        $ErrorActionPreference = 'Continue'
    }

    ##### FIRST CHECK AUTH PARAMETERS
    if ($PSBoundParameters.ContainsKey('Credential') -eq $true -and `
        -not [System.String]::IsNullOrEmpty($Credential))
    {
        if ($Credential.Username -notmatch ".onmicrosoft.")
        {
            Write-Warning -Message "We recommend providing the username in the format of <tenant>.onmicrosoft.* for the Credential property."
        }
    }

    if ($PSBoundParameters.ContainsKey('CertificatePath') -eq $true -and `
            $PSBoundParameters.ContainsKey('CertificatePassword') -eq $false)
    {
        throw 'You have to specify CertificatePassword when you specify CertificatePath'
    }

    if ($PSBoundParameters.ContainsKey('CertificatePassword') -eq $true -and `
            $PSBoundParameters.ContainsKey('CertificatePath') -eq $false)
    {
        throw 'You have to specify CertificatePath when you specify CertificatePassword'
    }

    if ($PSBoundParameters.ContainsKey('ApplicationId') -eq $true -and `
            $PSBoundParameters.ContainsKey('Credential') -eq $false -and `
            $PSBoundParameters.ContainsKey('TenantId') -eq $false)
    {
        throw 'You have to specify TenantId when you specify ApplicationId'
    }

    if ($PSBoundParameters.ContainsKey('ApplicationId') -eq $true -and `
            $PSBoundParameters.ContainsKey('TenantId') -eq $true -and `
            $PSBoundParameters.ContainsKey('Credential') -eq $false -and `
        ($PSBoundParameters.ContainsKey('CertificateThumbprint') -eq $false -and `
                $PSBoundParameters.ContainsKey('ApplicationSecret') -eq $false -and `
                $PSBoundParameters.ContainsKey('CertificatePath') -eq $false))
    {
        throw 'You have to specify ApplicationSecret, CertificateThumbprint or CertificatePath when you specify ApplicationId/TenantId'
    }

    if (($PSBoundParameters.ContainsKey('ApplicationId') -eq $false -or `
                $PSBoundParameters.ContainsKey('TenantId') -eq $false) -and `
        ($PSBoundParameters.ContainsKey('Credential') -eq $false -and `
                $PSBoundParameters.ContainsKey('CertificateThumbprint') -eq $true -or `
                $PSBoundParameters.ContainsKey('ApplicationSecret') -eq $true -or `
                $PSBoundParameters.ContainsKey('CertificatePath') -eq $true))
    {
        throw 'You have to specify ApplicationId and TenantId when you specify ApplicationSecret, CertificateThumbprint or CertificatePath'
    }

    # Default to Credential if no authentication mechanism were provided
    if ($PSBoundParameters.ContainsKey('Credential') -eq $false -and `
            $ManagedIdentity.IsPresent -eq $false -and `
            $PSBoundParameters.ContainsKey('ApplicationId') -eq $false -and `
            $PSBoundParameters.ContainsKey('AccessTokens') -eq $false)
    {
        $Credential = Get-Credential
    }

    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()

    $data.Add('Path', [System.String]::IsNullOrEmpty($Path))
    $data.Add('FileName', $null -ne [System.String]::IsNullOrEmpty($FileName))
    $data.Add('Components', $Components)
    $data.Add('Workloads', $Workloads)
    #endregion

    # Make sure we are not connected to Microsoft Graph on another tenant
    try
    {
        Confirm-M365DSCLoadedModule -ModuleName 'Microsoft.Graph.Authentication'
        if ($null -ne (Get-MgContext))
        {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            Reset-MSCloudLoginConnectionProfileContext -Workload 'MicrosoftGraph'
        }
    }
    catch
    {
        Write-Verbose -Message 'No existing connections to Microsoft Graph'
    }

    $Tenant = Get-M365DSCTenantNameFromParameterSet -ParameterSet $PSBoundParameters
    $ConnectionMode = Get-M365DSCAuthenticationMode $PSBoundParameters
    $data.Add('Tenant', $Tenant)
    $currentExportID = (New-Guid).ToString()
    $data.Add('M365DSCExportId', $currentExportID)
    $data.Add('ConnectionMode', $ConnectionMode)

    # Define connection to Graph parameters because it is required by the telemetry.
    if ($null -eq $Script:M365DSCTelemetryConnectionToGraphParams -or `
        ($null -ne $Script:M365DSCTelemetryConnectionToGraphParams -and `
        $Script:M365DSCTelemetryConnectionToGraphParams.Keys.Count -eq 0))
    {
        $Script:M365DSCTelemetryConnectionToGraphParams = @{
            Credential            = $Credential
            ApplicationId         = $ApplicationId
            TenantId              = $TenantId
            ApplicationSecret     = $ApplicationSecret
            CertificateThumbprint = $CertificateThumbprint
            CertificatePassword   = $CertificatePassword
            CertificatePath       = $CertificatePath
            Identity              = $ManagedIdentity.IsPresent
            AccessTokens          = $AccessTokens
        }
    }

    Add-M365DSCTelemetryEvent -Type 'ExportInitiated' -Data $data
    Initialize-M365DSCAllResourcesDictionary
    if ($PSBoundParameters.ContainsKey('TokenReplacement')) {
        Set-M365DSCStringReplacementMap -Map $TokenReplacement
    }

    $resourceSettings = Get-M365DSCResourceSettings
    if ($null -ne $Workloads)
    {
        Write-M365DSCHost -Message "Exporting Microsoft 365 configuration for Workloads: $($Workloads -join ', ')"
        Start-M365DSCConfigurationExtract -Credential $Credential `
            -Workloads $Workloads `
            -ExcludeComponents $ExcludeComponents `
            -Mode $Mode `
            -Path $Path -FileName $FileName `
            -ConfigurationName $ConfigurationName `
            -ApplicationId $ApplicationId `
            -ApplicationSecret $ApplicationSecret `
            -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint `
            -CertificatePath $CertificatePath `
            -CertificatePassword $CertificatePassword `
            -ManagedIdentity:$ManagedIdentity `
            -AccessTokens $AccessTokens `
            -GenerateInfo $GenerateInfo `
            -Filters $Filters `
            -Validate:$Validate `
            -Parallel:$Parallel `
            -ResourceSettings $resourceSettings `
            -ErrorAction $ErrorActionPreference `
            -WithStatistics:$WithStatistics
    }
    elseif ($null -ne $Components)
    {
        Write-M365DSCHost -Message "Exporting Microsoft 365 configuration for Components: $($Components -join ', ')"
        Start-M365DSCConfigurationExtract -Credential $Credential `
            -Components $Components `
            -ExcludeComponents $ExcludeComponents `
            -Path $Path -FileName $FileName `
            -ConfigurationName $ConfigurationName `
            -ApplicationId $ApplicationId `
            -ApplicationSecret $ApplicationSecret `
            -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint `
            -CertificatePath $CertificatePath `
            -CertificatePassword $CertificatePassword `
            -ManagedIdentity:$ManagedIdentity `
            -AccessTokens $AccessTokens `
            -GenerateInfo $GenerateInfo `
            -Filters $Filters `
            -Validate:$Validate `
            -Parallel:$Parallel `
            -ResourceSettings $resourceSettings `
            -ErrorAction $ErrorActionPreference `
            -WithStatistics:$WithStatistics
    }
    elseif ($null -ne $Mode)
    {
        Write-M365DSCHost -Message "Exporting Microsoft 365 configuration for Mode: $Mode"
        Start-M365DSCConfigurationExtract -Credential $Credential `
            -Mode $Mode `
            -ExcludeComponents $ExcludeComponents `
            -Path $Path -FileName $FileName `
            -ConfigurationName $ConfigurationName `
            -ApplicationId $ApplicationId `
            -ApplicationSecret $ApplicationSecret `
            -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint `
            -CertificatePath $CertificatePath `
            -CertificatePassword $CertificatePassword `
            -ManagedIdentity:$ManagedIdentity `
            -AccessTokens $AccessTokens `
            -GenerateInfo $GenerateInfo `
            -AllComponents `
            -Filters $Filters `
            -Validate:$Validate `
            -Parallel:$Parallel `
            -ResourceSettings $resourceSettings `
            -ErrorAction $ErrorActionPreference `
            -WithStatistics:$WithStatistics
    }

    # Clear the exported resource instances' names Global variable
    $Global:M365DSCExportedResourceInstancesNames = $null
    $Global:M365DSCExportInProgress = $false

    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    if ([System.String]::IsNullOrEmpty($data.Tenant) -and -not [System.String]::IsNullOrEmpty($TenantId))
    {
        $data.Add('Tenant', $TenantId)
    }
    else
    {
        $data.Add('Tenant', $Tenant)
    }
    $data.Add('M365DSCExportId', $currentExportID)
    $data.Add('ConnectionMode', $ConnectionMode)
    $timeTaken = [System.DateTime]::Now.Subtract($currentStartDateTime)
    $data.Add('TotalSeconds',$timeTaken.TotalSeconds)
    Add-M365DSCTelemetryEvent -Type 'ExportCompleted' -Data $data
}

<#
.DESCRIPTION
This function tests the code page of the current terminal session.

.EXAMPLE
Test-CodePage

.FUNCTIONALITY
Private
#>
function Test-CodePage
{
    if ([System.Text.Encoding]::Default.CodePage -ne 65001)
    {
        Write-Warning -Message 'The code page of the current session is not set to UTF-8. This may cause issues with Unicode characters.
         To change the code page to UTF-8, you have the following options:
         * Using the control panel: intl.cpl --> Administrative --> Change system locale --> Beta: Use Unicode UTF-8 for worldwide language support
         * Using PowerShell: Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "ACP" -Value 65001
         After that, you need to restart the PowerShell session.'
    }
}

<#
.Description
This function retrieves the various endpoint urls based on the cloud environment.

.Example
Get-M365DSCAPIEndpoint -TenantId 'contoso.onmicrosoft.com'

.Functionality
Private
#>
function Get-M365DSCAPIEndpoint
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $TenantId
    )

    try
    {
        $webrequest = Invoke-WebRequest -Uri "https://login.windows.net/$($TenantId)/.well-known/openid-configuration" -UseBasicParsing
        $response = ConvertFrom-Json $webrequest.Content
        $tenantRegionScope = $response."tenant_region_scope"

        $endpoints = @{
            AzureManagement = $null
        }

        switch ($tenantRegionScope)
        {
            'USGov'
            {
                if ($null -ne $response.'tenant_region_sub_scope' -and $response.'tenant_region_sub_scope' -eq 'DODCON')
                {
                    $endpoints.AzureManagement = "https://management.usgovcloudapi.net"
                }
            }
            default
            {
                $endpoints.AzureManagement = "https://management.azure.com"
            }
        }
        return $endpoints
    }
    catch
    {
        throw $_
    }
}

<#
.Description
This function gets the onmicrosoft.com name of the tenant

.Functionality
Internal
#>
function Get-M365DSCTenantDomain
{
    [CmdletBinding(DefaultParameterSetName = 'AppId')]
    [OutputType([System.String])]
    param
    (
        [Parameter(ParameterSetName = 'AppId', Mandatory = $true)]
        [System.String]
        $ApplicationId,

        [Parameter(Mandatory = $true)]
        [System.String]
        $TenantId,

        [Parameter(ParameterSetName = 'AppId')]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter(ParameterSetName = 'AppId')]
        [System.String]
        $CertificateThumbprint,

        [Parameter(ParameterSetName = 'AppId')]
        [System.String]
        $CertificatePath,

        [Parameter(ParameterSetName = 'MID')]
        [Switch]
        $ManagedIdentity
    )

    if ([System.String]::IsNullOrEmpty($CertificatePath))
    {
        $null = New-M365DSCConnection -Workload 'MicrosoftGraph' `
            -InboundParameters $PSBoundParameters

        try
        {
            $tenantDetails = (Invoke-MgGraphRequest -Uri "/beta/organization" -Method GET -ErrorAction 'Stop').value
            $defaultDomain = $tenantDetails.verifiedDomains | Where-Object -FilterScript { $_.isInitial }

            return $defaultDomain.name
        }
        catch
        {
            if ($_.Exception.Message -eq 'Insufficient privileges to complete the operation.')
            {
                New-M365DSCLogEntry `
                    -Message 'Error retrieving Organizational information: Missing Organization.Read.All permission. ' `
                    -Exception $_ `
                    -Source $($MyInvocation.MyCommand.Source) `
                    -TenantId $TenantId `
                    -Credential $Credential

                return [System.String]::Empty
            }

            throw $_
        }
    }

    if ($TenantId.Contains('onmicrosoft'))
    {
        return $TenantId
    }
    else
    {
        throw 'TenantID must be in format contoso.onmicrosoft.com'
    }
}

<#
.Description
This function gets the DNS domain used in the specified credential

.Functionality
Internal
#>
function Get-M365DSCOrganization
{
    param
    (
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $TenantId
    )

    if ($null -ne $Credential -and $Credential.UserName.Contains('@'))
    {
        $organization = $Credential.UserName.Split('@')[1]
        return $organization
    }

    if (-not [System.String]::IsNullOrEmpty($TenantId))
    {
        if ($TenantId.Contains('.'))
        {
            $organization = $TenantId
            return $organization
        }
        else
        {
            Throw 'Tenant ID must be name of the tenant, e.g. contoso.onmicrosoft.com'
        }

    }
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
.Description
This function creates a new connection to the specified M365 workload

.Functionality
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
                $invalid = $_.Credential.Username -notmatch ".onmicrosoft."
                if (-not $invalid)
                {
                    return $true
                }
                else
                {
                    Write-Warning -Message "We recommend providing the username in the format of <tenant>.onmicrosoft.* for the Credential property."
                }
            }

            if ($null -ne $_.TenantId)
            {
                $parseGuid = [System.Guid]::Empty
                $isValid = [System.Guid]::TryParse($_.TenantId, [ref]$parseGuid)
                if ($isValid)
                {
                    throw "Please provide the tenant name (e.g., contoso.onmicrosoft.com) for TenantId instead of its GUID."
                }

                $isValid = $_.TenantId -match ".onmicrosoft."
                if ($isValid)
                {
                    return $true
                }
                else
                {
                    Write-Warning -Message "We recommend providing the tenant name in format <tenant>.onmicrosoft.* for TenantId."
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

    $requiredModules = Get-M365DSCRequiredModules
    foreach ($requiredModule in $requiredModules)
    {
        Write-Verbose -Message "Ensuring required module '$requiredModule' is loaded."
        Confirm-M365DSCLoadedModule -ModuleName $requiredModule
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
        Write-Verbose -Message "Initializing the Connected To Workloads List."
        $Script:M365ConnectedToWorkloads = @()
    }

    # Get the ApplicationSecret parameter back as a string.
    if ($InboundParameters.ApplicationSecret)
    {
        $InboundParameters.ApplicationSecret = $InboundParameters.ApplicationSecret.GetNetworkCredential().Password
        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationSecret'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationSecret', $InboundParameters.ApplicationSecret)
        }
    }

    # Case both authentication methods are attempted
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
    # Case no authentication method is specified
    elseif ($null -eq $InboundParameters.Credential -and `
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
    # Case only Credential is specified
    elseif ($null -ne $InboundParameters.Credential -and `
            [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint))
    {
        Write-Verbose -Message 'Credential was specified. Connecting via User Principal'
        if ([System.String]::IsNullOrEmpty($Url))
        {
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('Credential'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('Credential', $InboundParameters.Credential)
            }
            Connect-M365Tenant -Workload $Workload `
                -Credential $InboundParameters.Credential `
                -EnableSearchOnlySession:$EnableSearchOnlySession

            if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-Credential")
            {
                $data.Add('ConnectionMode', 'Credentials')
                try
                {
                    if (-not $Data.ContainsKey('Tenant'))
                    {
                        $tenantId = $InboundParameters.Credential.Username.Split('@')[1]
                        $data.Add('Tenant', $tenantId)

                        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
                        {
                            $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $tenantId)
                        }
                    }
                }
                catch
                {
                    Write-Verbose -Message $_
                }

                Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
                $Script:M365ConnectedToWorkloads += "$Workload-Credential"
            }
            return 'Credentials'
        }
        if ($InboundParameters.ContainsKey('Credential') -and
            $null -ne $InboundParameters.Credential)
        {
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('Credential'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('Credential', $InboundParameters.Credential)
            }
            Connect-M365Tenant -Workload $Workload `
                -Credential $InboundParameters.Credential `
                -Url $Url `
                -EnableSearchOnlySession:$EnableSearchOnlySession

            if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-Credential")
            {
                $data.Add('ConnectionMode', 'Credential')
                try
                {
                    if (-not $Data.ContainsKey('Tenant'))
                    {
                        $tenantId = $InboundParameters.Credential.Username.Split('@')[1]
                        $data.Add('Tenant', $tenantId)
                        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
                        {
                            $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $tenantId)
                        }
                    }
                }
                catch
                {
                    Write-Verbose -Message $_
                }

                Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
                $Script:M365ConnectedToWorkloads += "$Workload-Credential"
            }
            return 'Credentials'
        }
    }
    # Case only Credential with ApplicationId is specified
    elseif ($null -ne $InboundParameters.Credential -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint))
    {
        if ([System.String]::IsNullOrEmpty($Url))
        {
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('Credential'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('Credential', $InboundParameters.Credential)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationId', $InboundParameters.ApplicationId)
            }
            Connect-M365Tenant -Workload $Workload `
                -ApplicationId $InboundParameters.ApplicationId `
                -Credential $InboundParameters.Credential `
                -EnableSearchOnlySession:$EnableSearchOnlySession

            if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-CredentialsWithApplicationId")
            {
                $data.Add('ConnectionMode', 'CredentialsWithApplicationId')

                try
                {
                    if (-not $Data.ContainsKey('Tenant'))
                    {
                        $tenantId = $InboundParameters.Credential.Username.Split('@')[1]
                        $data.Add('Tenant', $tenantId)
                        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
                        {
                            $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $tenantId)
                        }
                    }
                }
                catch
                {
                    Write-Verbose -Message $_
                }

                Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
                $Script:M365ConnectedToWorkloads += "$Workload-CredentialsWithApplicationId"
            }
            return 'CredentialsWithApplicationId'
        }
        else
        {
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationId', $InboundParameters.ApplicationId)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('Credential'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('Credential', $InboundParameters.Credential)
            }
            Connect-M365Tenant -Workload $Workload `
                -ApplicationId $InboundParameters.ApplicationId `
                -Credential $InboundParameters.Credential `
                -Url $Url `
                -EnableSearchOnlySession:$EnableSearchOnlySession

            if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-CredentialsWithApplicationId")
            {
                $data.Add('ConnectionMode', 'CredentialsWithApplicationId')

                try
                {
                    if (-not $Data.ContainsKey('Tenant'))
                    {
                        $tenantId = $InboundParameters.Credential.Username.Split('@')[1]
                        $data.Add('Tenant', $tenantId)
                        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationId'))
                        {
                            $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationId', $tenantId)
                        }
                    }
                }
                catch
                {
                    Write-Verbose -Message $_
                }

                Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
                $Script:M365ConnectedToWorkloads += "$Workload-CredentialsWithApplicationId"
            }
            return 'CredentialsWithApplicationId'
        }
    }
    # Case only the ServicePrincipal with CertificatePath parameters are specified
    elseif ($null -eq $InboundParameters.Credential -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.CertificatePath) -and `
            $null -ne $InboundParameters.CertificatePassword)
    {
        if ([System.String]::IsNullOrEmpty($url))
        {
            Write-Verbose -Message 'ApplicationId, TenantId, CertificatePath & CertificatePassword were specified. Connecting via Service Principal'

            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationId', $InboundParameters.ApplicationId)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('CertificatePassword'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('CertificatePassword', $InboundParameters.CertificatePassword.Password)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('CertificatePath'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('CertificatePath', $InboundParameters.CertificatePath)
            }
            Connect-M365Tenant -Workload $Workload `
                -ApplicationId $InboundParameters.ApplicationId `
                -TenantId $InboundParameters.TenantId `
                -CertificatePassword $InboundParameters.CertificatePassword.Password `
                -CertificatePath $InboundParameters.CertificatePath `
                -EnableSearchOnlySession:$EnableSearchOnlySession

            if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-ServicePrincipalWithPath")
            {
                $data.Add('ConnectionMode', 'ServicePrincipalWithPath')
                if (-not $data.ContainsKey('Tenant'))
                {
                    $data.Add('Tenant', $InboundParameters.TenantId)
                    if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
                    {
                        $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
                    }
                }
                Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
                $Script:M365ConnectedToWorkloads += "$Workload-ServicePrincipalWithPath"
            }
            return 'ServicePrincipalWithPath'
        }
        #endregion

        # Case no authentication method is specified
        if ($null -eq $InboundParameters.Credential -and `
                [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
                [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
                [System.String]::IsNullOrEmpty($InboundParameters.CertificateThumbprint))
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
        else
        {
            $data.Add('ConnectionMode', 'ServicePrincipalWithPath')
            if (-not $data.ContainsKey('Tenant'))
            {
                if (-not [System.String]::IsNullOrEmpty($InboundParameters.TenantId))
                {
                    $data.Add('Tenant', $InboundParameters.TenantId)
                }
                elseif ($ null -ne $InboundParameters.Credential)
                {
                    $data.Add('Tenant', $InboundParameters.Credential.Split('@')[1])
                }
            }
            Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'

            return 'ServicePrincipalWithPath'
        }
    }
    # Case only the ApplicationSecret, TenantId and ApplicationID are specified
    elseif ($null -eq $InboundParameters.Credential -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.ApplicationId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.TenantId) -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.ApplicationSecret))
    {
        if ([System.String]::IsNullOrEmpty($url))
        {
            Write-Verbose -Message 'ApplicationId, TenantId, ApplicationSecret were specified. Connecting via Service Principal'

            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationId', $InboundParameters.ApplicationId)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationSecret'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationSecret', $InboundParameters.ApplicationSecret)
            }
            Connect-M365Tenant -Workload $Workload `
                -ApplicationId $InboundParameters.ApplicationId `
                -TenantId $InboundParameters.TenantId `
                -ApplicationSecret $InboundParameters.ApplicationSecret `
                -EnableSearchOnlySession:$EnableSearchOnlySession

            if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-ServicePrincipalWithSecret")
            {
                $data.Add('ConnectionMode', 'ServicePrincipalWithSecret')
                if (-not $data.ContainsKey('Tenant'))
                {
                    $data.Add('Tenant', $InboundParameters.TenantId)
                }
                Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
                $Script:M365ConnectedToWorkloads += "$Workload-ServicePrincipalWithSecret"
            }
            return 'ServicePrincipalWithSecret'
        }
        else
        {
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationId', $InboundParameters.ApplicationId)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
            }
            if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationSecret'))
            {
                $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationSecret', $InboundParameters.ApplicationSecret)
            }
            Connect-M365Tenant -Workload $Workload `
                -ApplicationId $InboundParameters.ApplicationId `
                -TenantId $InboundParameters.TenantId `
                -ApplicationSecret $InboundParameters.ApplicationSecret `
                -Url $Url `
                -EnableSearchOnlySession:$EnableSearchOnlySession

            if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-ServicePrincipalWithSecret")
            {
                $data.Add('ConnectionMode', 'ServicePrincipalWithSecret')
                if (-not $data.ContainsKey('Tenant'))
                {
                    $data.Add('Tenant', $InboundParameters.TenantId)
                }
                Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
                $Script:M365ConnectedToWorkloads += "$Workload-ServicePrincipalWithSecret"
            }
            return 'ServicePrincipalWithSecret'
        }
    }
    elseif ($InboundParameters.CertificateThumbprint -and $InboundParameters.ApplicationId -and $InboundParameters.TenantId)
    {
        Write-Verbose -Message 'ApplicationId, TenantId, CertificateThumbprint were specified. Connecting via Service Principal'

        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('ApplicationId'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('ApplicationId', $InboundParameters.ApplicationId)
        }
        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
        }
        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('CertificateThumbprint'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('CertificateThumbprint', $InboundParameters.CertificateThumbprint)
        }
        Write-Verbose -Message "Calling into Connect-M365Tenant"
        Connect-M365Tenant -Workload $Workload `
            -ApplicationId $InboundParameters.ApplicationId `
            -TenantId $InboundParameters.TenantId `
            -CertificateThumbprint $InboundParameters.CertificateThumbprint `
            -Url $Url `
            -EnableSearchOnlySession:$EnableSearchOnlySession

        Write-Verbose -Message "Connection initiated."
        if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-ServicePrincipalWithThumbprint")
        {
            $data.Add('ConnectionMode', 'ServicePrincipalWithThumbprint')
            if (-not $data.ContainsKey('Tenant'))
            {
                $data.Add('Tenant', $InboundParameters.TenantId)
            }
            Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
            $Script:M365ConnectedToWorkloads += "$Workload-ServicePrincipalWithThumbprint"
        }
        return 'ServicePrincipalWithThumbprint'
    }
    # Case only the TenantId and Credentials parameters are specified
    elseif ($null -ne $InboundParameters.Credential -and `
    -not [System.String]::IsNullOrEmpty($InboundParameters.TenantId))
    {
        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('Credential'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('Credential', $InboundParameters.Credential)
        }
        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
        }
        Connect-M365Tenant -Workload $Workload `
            -TenantId $InboundParameters.TenantId `
            -Credential $InboundParameters.Credential `
            -Url $Url `
            -EnableSearchOnlySession:$EnableSearchOnlySession

        if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-CredentialsWithTenantId")
        {
            $data.Add('ConnectionMode', 'CredentialsWithTenantId')
            if (-not $data.ContainsKey('Tenant'))
            {
                $data.Add('Tenant', $InboundParameters.TenantId)
            }
            Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
            $Script:M365ConnectedToWorkloads += "$Workload-CredentialsWithTenantId"
        }
        return "CredentialsWithTenantId"
    }
    # Case only Managed Identity and TenantId are specified
    elseif ($InboundParameters.ManagedIdentity -and `
            -not [System.String]::IsNullOrEmpty($InboundParameters.TenantId))
    {
        Write-Verbose -Message 'Connecting via managed identity'

        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('Identity'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('Identity', $true)
        }
        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
        }
        Connect-M365Tenant -Workload $Workload `
            -Identity `
            -TenantId $InboundParameters.TenantId `
            -EnableSearchOnlySession:$EnableSearchOnlySession

        if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-ManagedIdentity")
        {
            $data.Add('ConnectionMode', 'ManagedIdentity')
            if (-not $data.ContainsKey('Tenant'))
            {
                $data.Add('Tenant', $InboundParameters.TenantId)
            }
            Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
            $Script:M365ConnectedToWorkloads += "$Workload-ManagedIdentity"
        }
        return 'ManagedIdentity'
    }
    # Case Access Token is Specified
    elseif ($null -ne $InboundParameters.AccessTokens -and `
    -not [System.String]::IsNullOrEmpty($InboundParameters.TenantId))
    {
        Write-Verbose -Message 'Connecting via Access Tokens'

        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('AccessTokens'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('AccessTokens', $InboundParameters.AccessTokens)
        }

        if (-not $Script:M365DSCTelemetryConnectionToGraphParams.ContainsKey('TenantId'))
        {
            $Script:M365DSCTelemetryConnectionToGraphParams.Add('TenantId', $InboundParameters.TenantId)
        }
        Connect-M365Tenant -Workload $Workload `
            -AccessTokens $InboundParameters.AccessTokens `
            -TenantId $InboundParameters.TenantId `
            -EnableSearchOnlySession:$EnableSearchOnlySession

        if (-not $Script:M365ConnectedToWorkloads -contains "$Workload-AccessTokens")
        {
            $data.Add('ConnectionMode', 'AccessTokens')
            if (-not $data.ContainsKey('Tenant'))
            {
                $data.Add('Tenant', $InboundParameters.TenantId)
            }
            Add-M365DSCTelemetryEvent -Data $data -Type 'Connection'
            $Script:M365ConnectedToWorkloads += "$Workload-AccessTokens"
        }
        return 'AccessTokens'
    }
    else
    {
        throw 'Could not determine authentication method'
    }
}

<#
.Description
This function gets the URL of the SPO Administration site

.Functionality
Internal
#>
function Get-SPOAdministrationUrl
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [switch]
        $UseMFA,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    $UseMFASwitch = @{}
    if ($UseMFA)
    {
        $UseMFASwitch.Add('UseMFA', $true)
    }

    Write-Verbose -Message 'Connection to Azure AD is required to automatically determine SharePoint Online admin URL...'
    $null = New-M365DSCConnection -Workload 'MicrosoftGraph' `
        -InboundParameters $PSBoundParameters
    Write-Verbose -Message 'Getting SharePoint Online admin URL...'
    $domain = Invoke-MgGraphRequest -Uri "beta/domains" -Method GET
    [Array]$defaultDomain = $domain | Where-Object { ($_.id -like '*.onmicrosoft.com' -or $_.id -like '*.onmicrosoft.de' -or $_.id -like '*.onmicrosoft.us') -and $_.isInitial -eq $true } # We don't use IsDefault here because the default could be a custom domain

    if ($defaultDomain[0].id -like '*.onmicrosoft.com*')
    {
        $global:tenantName = $defaultDomain[0].id -replace '.onmicrosoft.com', ''
    }
    elseif ($defaultDomain[0].id -like '*.onmicrosoft.de*')
    {
        $global:tenantName = $defaultDomain[0].id -replace '.onmicrosoft.de', ''
    }
    $global:AdminUrl = "https://$global:tenantName-admin.sharepoint.com"
    Write-Verbose -Message "SharePoint Online admin URL is $global:AdminUrl"
    return $global:AdminUrl
}

<#
.Description
This function gets the name of the M365 tenant

.Functionality
Internal
#>
function Get-M365TenantName
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [switch]
        $UseMFA,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    $UseMFASwitch = @{}
    if ($UseMFA)
    {
        $UseMFASwitch.Add('UseMFA', $true)
    }

    Write-Verbose -Message 'Connection to Azure AD is required to automatically determine SharePoint Online admin URL...'
    $null = New-M365DSCConnection -Workload 'MicrosoftGraph' `
        -InboundParameters $PSBoundParameters
    Write-Verbose -Message 'Getting SharePoint Online admin URL...'
    $domain = Invoke-MgGraphRequest -Uri "beta/domains" -Method GET
    [Array]$defaultDomain = $domain | Where-Object { ($_.id -like '*.onmicrosoft.com' -or $_.id -like '*.onmicrosoft.de') -and $_.isInitial -eq $true } # We don't use IsDefault here because the default could be a custom domain

    if ($defaultDomain[0].id -like '*.onmicrosoft.com*')
    {
        $tenantName = $defaultDomain[0].id -replace '.onmicrosoft.com', ''
    }
    elseif ($defaultDomain[0].id -like '*.onmicrosoft.de*')
    {
        $tenantName = $defaultDomain[0].id -replace '.onmicrosoft.de', ''
    }

    Write-Verbose -Message "M365 tenant name is $tenantName"
    return $tenantName
}

<#
.Description
This function downloads and installs the Dev branch of Microsoft365DSC on the local machine

.Parameter Scope
Specifies the scope of the update of the module. The default value is AllUsers(needs to run as elevated user).

.Example
Install-M365DSCDevBranch

.Example
Install-M365DSCDevBranch -Scope CurrentUser

.Functionality
Public
#>
function Install-M365DSCDevBranch
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet("CurrentUser", "AllUsers")]
        $Scope = "AllUsers"
    )

    try {

        $longPathsEnabled = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem').LongPathsEnabled -eq 1
        if (-not $longPathsEnabled)
        {
            $message = "Long paths are not enabled on this system. You may encounter issues with the installation of Microsoft365DSC because of long file names."
            $message += "To enable long paths, set the registry LongPathsEnabled DWORD entry to 1 in HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem."
            Write-Warning -Message $message
        }

        #region Download and Extract Dev branch's ZIP
        Write-Host 'Downloading the Zip package...' -NoNewline
        $url = 'https://github.com/microsoft/Microsoft365DSC/archive/Dev.zip'
        $output = "$($env:Temp)\dev.zip"
        $extractPath = $env:Temp + '\O365Dev'
        Write-Host 'Done' -ForegroundColor Green

        Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing

        Expand-Archive $output -DestinationPath $extractPath -Force
        #endregion

        #region Install All Dependencies
        $manifest = Import-PowerShellDataFile "$extractPath\Microsoft365DSC-Dev\Modules\Microsoft365DSC\Microsoft365DSC.psd1"
        $dependencies = $manifest.RequiredModules
        if ((-not(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) -and ($Scope -eq "AllUsers"))
        {
            Write-Error 'Cannot update the dependencies for Microsoft365DSC. You need to run this command as a local administrator.'
        }
        else
        {
            foreach ($dependency in $dependencies)
            {
                Write-Host "Installing {$($dependency.ModuleName)}..." -NoNewline
                $existingModule = Get-Module $dependency.ModuleName -ListAvailable | Where-Object -FilterScript { $_.Version -eq $dependency.RequiredVersion }
                if ($null -eq $existingModule)
                {
                    Install-Module $dependency.ModuleName -RequiredVersion $dependency.RequiredVersion -Force -AllowClobber -Scope $Scope | Out-Null
                }
                Import-Module $dependency.ModuleName -Force | Out-Null
                Write-Host 'Done' -ForegroundColor Green
            }
        }
        #endregion

        #region Install M365DSC
        Write-Host 'Updating the Core Microsoft365DSC module...' -NoNewline
        $defaultPath = 'C:\Program Files\WindowsPowerShell\Modules\Microsoft365DSC\'
        $currentVersionPath = $defaultPath + ([Version]$($manifest.ModuleVersion)).ToString()

        Copy-Item "$extractPath\Microsoft365DSC-Dev\Modules\Microsoft365DSC\*" `
            -Destination $defaultPath -Recurse -Force

        Import-Module ($defaultPath + 'Microsoft365DSC.psd1') -Force | Out-Null
        $oldModule = Get-Module 'Microsoft365DSC' | Where-Object -FilterScript { $_.ModuleBase -eq $currentVersionPath }
        Remove-Module $oldModule -Force | Out-Null
        if (Test-Path $currentVersionPath)
        {
            try
            {
                Remove-Item $currentVersionPath -Recurse -Confirm:$false -Force `
                    -ErrorAction Stop
            }
            catch
            {
                Write-Verbose -Message $_
            }
        }
        Write-Host 'Done' -ForegroundColor Green
        #endregion
    }
    catch
    {
        New-M365DSCLogEntry -Message 'Error installing Dev Branch:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source)
        Write-Error $_
    }
}

<#
.Description
This function downloads all apps installed in SPO

.Functionality
Internal
#>
function Get-AllSPOPackages
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable[]])]
    param
    (
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
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity
    )

    try
    {
        $null = New-M365DSCConnection -Workload 'PnP' `
            -InboundParameters $PSBoundParameters

        $tenantAppCatalogUrl = Get-PnPTenantAppCatalogUrl -ErrorAction Stop

        $null = New-M365DSCConnection -Workload 'PnP' `
            -InboundParameters $PSBoundParameters `
            -Url $tenantAppCatalogUrl

        $filesToDownload = @()
        $allFiles = @()
        if ($null -ne $tenantAppCatalogUrl)
        {
            try
            {
                [Array]$spfxFiles = @(Find-PnPFile -List 'AppCatalog' -Match '*.sppkg' -ErrorAction Stop)
                [Array]$appFiles = @(Find-PnPFile -List 'AppCatalog' -Match '*.app' -ErrorAction Stop)

                $allFiles = $spfxFiles + $appFiles

                foreach ($file in $allFiles)
                {
                    $filesToDownload += @{
                        Name = $file.Name
                        Site = $tenantAppCatalogUrl
                        Title = $file.Title
                    }
                }
            }
            catch
            {
                New-M365DSCLogEntry -Message $_.Exception.Message `
                -Exception $_ `
                -Source $($MyInvocation.MyCommand.Source) `
                -TenantId $TenantId `
                -Credential $Credential
            }
        }

        return $filesToDownload
    }
    catch
    {
        Write-Verbose -Message $_
    }

    return $null
}

<#
.Description
This function removes all items that have a Null value from the provided hashtable

.Functionality
Internal
#>
function Remove-NullEntriesFromHashtable
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.COllections.HashTable]
        $Hash
    )

    $keysToRemove = @()
    foreach ($key in $Hash.Keys)
    {
        if ([System.String]::IsNullOrEmpty($Hash.$key))
        {
            $keysToRemove += $key
        }
    }

    foreach ($key in $keysToRemove)
    {
        $Hash.Remove($key) | Out-Null
    }

    return $Hash
}

<#
.Description
This function compares a created export with the specified M365DSC Blueprint

.Parameter BluePrintUrl
Specifies the url to the blueprint to which the tenant should be compared.

.Parameter OutputReportPath
Specifies the path of the report that will be created.

.Parameter Credentials
Specifies the credentials that will be used for authentication.

.Parameter ApplicationId
Specifies the application id to be used for authentication.

.Parameter ApplicationSecret
Specifies the application secret of the application to be used for authentication.

.Parameter TenantId
Specifies the id of the tenant.

.Parameter CertificateThumbprint
Specifies the thumbprint to be used for authentication.

.Parameter CertificatePassword
Specifies the password of the PFX file which is used for authentication.

.Parameter CertificatePath
Specifies the path of the PFX file which is used for authentication.

.Parameter HeaderFilePath
Specifies that file that contains a custom header for the report.

.Parameter ExcludedProperties
Specifies the name of parameters that should not be assessed as part of the report. The names speficied will apply to all resources where they are encountered.

.Parameter ExcludedResources
Specifies the name of resources that should not be assessed as part of the report.

.Parameter DriftOnly
Specifies if the report should only show properties drifts and not missing instances.

.Example
Assert-M365DSCBlueprint -BluePrintUrl 'C:\DS\blueprint.m365' -OutputReportPath 'C:\DSC\BlueprintReport.html'

.Example
Assert-M365DSCBlueprint -BluePrintUrl 'C:\DS\blueprint.m365' -OutputReportPath 'C:\DSC\BlueprintReport.html' -Credentials $credentials -HeaderFilePath 'C:\DSC\ReportCustomHeader.html'

.Example
Assert-M365DSCBlueprint -BluePrintUrl 'C:\DS\blueprint.m365' -OutputReportPath 'C:\DSC\BlueprintReport.html' -ApplicationId $clientid -TenantId $tenantId -CertificateThumbprint $certthumbprint -HeaderFilePath 'C:\DSC\ReportCustomHeader.html'

.Functionality
Public
#>
function Assert-M365DSCBlueprint
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $BluePrintUrl,

        [Parameter(Mandatory = $true)]
        [System.String]
        $OutputReportPath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credentials,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $HeaderFilePath,

        [Parameter()]
        [System.String]
        [ValidateSet('HTML', 'JSON')]
        $Type = 'HTML',

        [Parameter()]
        [System.String[]]
        $ExcludedProperties,

        [Parameter()]
        [System.String[]]
        $ExcludedResources,

        [Parameter()]
        [System.Boolean]
        $DriftOnly = $true
    )

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add('Event', 'AssertBlueprint')
    $data.Add('BluePrint', $BluePrintUrl)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $TempBluePrintName = 'TempBlueprint_' + (New-Guid).ToString() + '.M365'
    $LocalBluePrintPath = Join-Path -Path $env:Temp -ChildPath $TempBluePrintName
    try
    {
        # Download the BluePrint locally in a temp location
        Invoke-WebRequest -Uri $BluePrintUrl -OutFile $LocalBluePrintPath -UseBasicParsing
    }
    catch
    {
        # If the download failed, we assume the provided Url was a local path
        # and we try copying the item instead.
        try
        {
            Copy-Item -Path $BluePrintUrl -Destination $LocalBluePrintPath
        }
        catch
        {
            throw $_
        }
    }

    if (Test-Path -Path $LocalBluePrintPath)
    {
        # Parse the content of the BluePrint into an array of PowerShell Objects
        $fileContent = Get-Content $LocalBluePrintPath -Raw
        $startPosition = $fileContent.IndexOf(' -ModuleVersion')
        if ($startPosition -gt 0)
        {
            $endPosition = $fileContent.IndexOf("`r", $startPosition)
            $fileContent = $fileContent.Remove($startPosition, $endPosition - $startPosition)
        }

        try
        {
            $parsedBluePrint = ConvertTo-DSCObject -Content $fileContent
        }
        catch
        {
            throw $_
        }

        # Generate an Array of Resource Types contained in the BluePrint
        $ResourcesInBluePrint = @()
        foreach ($resource in $parsedBluePrint)
        {
            if ($resource.ResourceName -in $ExcludedResources)
            {
                continue
            }
            if ($ResourcesInBluePrint -notcontains $resource.ResourceName)
            {
                $ResourcesInBluePrint += $resource.ResourceName
            }
        }

        if ([String]::IsNullOrEmpty($ResourcesInBluePrint))
        {
            if (![String]::IsNullOrEmpty($ExcludedResources))
            {
                Write-Host 'All resources were excluded from BluePrint, aborting'
            }
            else
            {
                Write-Host 'Malformed BluePrint, aborting'
            }
            break
        }

        Write-Host "Selected BluePrint contains ($($ResourcesInBluePrint.Length)) components to assess."

        # Call the Export-M365DSCConfiguration cmdlet to extract only the resource
        # types contained within the BluePrint;
        Write-Host "Initiating the Export of those ($($ResourcesInBluePrint.Length)) components from the tenant..."
        $TempExportName = 'TempExport_' + (New-Guid).ToString() + '.ps1'
        Export-M365DSCConfiguration -Components $ResourcesInBluePrint `
            -Path $env:temp `
            -FileName $TempExportName `
            -Credential $Credentials `
            -ApplicationId $ApplicationId `
            -ApplicationSecret $ApplicationSecret `
            -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint `
            -CertificatePath $CertificatePath `
            -CertificatePassword $CertificatePassword

        # Call the New-M365DSCDeltaReport configuration to generate the Delta Report between
        # the BluePrint and the extracted resources;
        $ExportPath = Join-Path -Path $env:Temp -ChildPath $TempExportName
        New-M365DSCDeltaReport -Source $ExportPath `
            -Destination $LocalBluePrintPath `
            -OutputPath $OutputReportPath `
            -DriftOnly $DriftOnly `
            -IsBlueprintAssessment:$true `
            -HeaderFilePath $HeaderFilePath `
            -Type $Type `
            -ExcludedProperties $ExcludedProperties `
            -ExcludedResources $ExcludedResources

        # Clean up the temporary files
        Remove-Item $LocalBluePrintPath -Force -ErrorAction SilentlyContinue
        Remove-Item $ExportPath -Force -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Error "M365DSC Template Path {$LocalBluePrintPath} does not exist."
    }
}

<#
.Description
This function updates the exported results with the specified authentication method

.Functionality
Internal
#>
function Update-M365DSCExportAuthenticationResults
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        [ValidateSet('ServicePrincipalWithThumbprint', 'ServicePrincipalWithSecret', 'ServicePrincipalWithPath', 'CredentialsWithTenantId', 'CredentialsWithApplicationId', 'Credentials', 'ManagedIdentity', 'AccessTokens')]
        $ConnectionMode,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Results
    )

    $noEscape = @()
    if ($Results.ContainsKey('ManagedIdentity') -and -not $Results.ManagedIdentity)
    {
        $Results.Remove('ManagedIdentity')
    }
    if ($ConnectionMode -eq 'Credentials')
    {
        $Results.Credential = Resolve-Credentials -UserName 'credential'
        $noEscape += 'Credential'
        if ($Results.ContainsKey('ApplicationId'))
        {
            $Results.Remove('ApplicationId') | Out-Null
        }
        if ($Results.ContainsKey('TenantId'))
        {
            $Results.Remove('TenantId') | Out-Null
        }
        if ($Results.ContainsKey('ApplicationSecret'))
        {
            $Results.Remove('ApplicationSecret') | Out-Null
        }
        if ($Results.ContainsKey('CertificateThumbprint'))
        {
            $Results.Remove('CertificateThumbprint') | Out-Null
        }
        if ($Results.ContainsKey('CertificatePath'))
        {
            $Results.Remove('CertificatePath') | Out-Null
        }
        if ($Results.ContainsKey('CertificatePassword'))
        {
            $Results.Remove('CertificatePassword') | Out-Null
        }
    }
    elseif ($ConnectionMode -eq 'CredentialsWithTenantId')
    {
        $Results.Credential = Resolve-Credentials -UserName 'credential'
        $noEscape += 'Credential'
        if ($Results.ContainsKey('ApplicationId'))
        {
            $Results.Remove('ApplicationId') | Out-Null
        }
        if ($Results.ContainsKey('ApplicationSecret'))
        {
            $Results.Remove('ApplicationSecret') | Out-Null
        }
        if ($Results.ContainsKey('CertificateThumbprint'))
        {
            $Results.Remove('CertificateThumbprint') | Out-Null
        }
        if ($Results.ContainsKey('CertificatePath'))
        {
            $Results.Remove('CertificatePath') | Out-Null
        }
        if ($Results.ContainsKey('CertificatePassword'))
        {
            $Results.Remove('CertificatePassword') | Out-Null
        }
    }
    else
    {
        if ($Results.ContainsKey('Credential') -and $ConnectionMode -ne 'CredentialsWithApplicationId')
        {
            $Results.Remove('Credential') | Out-Null
        }
        elseif ($Results.ContainsKey('Credential') -and $ConnectionMode -eq 'CredentialsWithApplicationId')
        {
            $Results.Credential = Resolve-Credentials -UserName 'credential'
            $noEscape += 'Credential'
        }
        if (-not [System.String]::IsNullOrEmpty($Results.ApplicationId))
        {
            $Results.ApplicationId = "`$ConfigurationData.NonNodeData.ApplicationId"
            $noEscape += 'ApplicationId'
        }
        else
        {
            try
            {
                $Results.Remove('ApplicationId') | Out-Null
            }
            catch
            {
                Write-Verbose -Message 'Error removing ApplicationId from Update-M365DSCExportAuthenticationResults'
            }
        }
        if (-not [System.String]::IsNullOrEmpty($Results.CertificateThumbprint))
        {
            $Results.CertificateThumbprint = "`$ConfigurationData.NonNodeData.CertificateThumbprint"
            $noEscape += 'CertificateThumbprint'
        }
        else
        {
            try
            {
                $Results.Remove('CertificateThumbprint') | Out-Null
            }
            catch
            {
                Write-Verbose -Message 'Error removing CertificateThumbprint from Update-M365DSCExportAuthenticationResults'
            }
        }
        if (-not [System.String]::IsNullOrEmpty($Results.CertificatePath))
        {
            $Results.CertificatePath = "`$ConfigurationData.NonNodeData.CertificatePath"
            $noEscape += 'CertificatePath'
        }
        else
        {
            try
            {
                $Results.Remove('CertificatePath') | Out-Null
            }
            catch
            {
                Write-Verbose -Message 'Error removing CertificatePath from Update-M365DSCExportAuthenticationResults'
            }
        }
        if (-not [System.String]::IsNullOrEmpty($Results.TenantId))
        {
            $Results.TenantId = "`$ConfigurationData.NonNodeData.TenantId"
            $noEscape += 'TenantId'
        }
        else
        {
            try
            {
                $Results.Remove('TenantId') | Out-Null
            }
            catch
            {
                Write-Verbose -Message 'Error removing TenantId from Update-M365DSCExportAuthenticationResults'
            }
        }
        if (-not [System.String]::IsNullOrEmpty($Results.ApplicationSecret))
        {
            $Results.ApplicationSecret = "New-Object System.Management.Automation.PSCredential ('ApplicationSecret', (ConvertTo-SecureString `$ConfigurationData.NonNodeData.ApplicationSecret -AsPlainText -Force))"
            $noEscape += 'ApplicationSecret'
        }
        else
        {
            try
            {
                $Results.Remove('ApplicationSecret') | Out-Null
            }
            catch
            {
                Write-Verbose -Message 'Error removing ApplicationSecret from Update-M365DSCExportAuthenticationResults'
            }
        }
        if ($null -ne $Results.CertificatePassword)
        {
            $Results.CertificatePassword = Resolve-Credentials -UserName 'CertificatePassword'
        }
        else
        {
            try
            {
                $Results.Remove('CertificatePassword') | Out-Null
            }
            catch
            {
                Write-Verbose -Message 'Error removing CertificatePassword from Update-M365DSCExportAuthenticationResults'
            }
        }

        if ($null -ne $Results.AccessTokens)
        {
            $results.AccessTokens = "`$ConfigurationData.NonNodeData.AccessTokens"
            $noEscape += 'AccessTokens'
        }
    }

    return @{
        Results = $Results
        NoEscape = $noEscape
    }
}

<#
.DESCRIPTION
    This function sets the string replacement map used during export.

.FUNCTIONALITY
    Internal
#>
function Set-M365DSCStringReplacementMap
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Collections.Hashtable]
        $Map,

        [Parameter()]
        [switch]
        $Clear
    )

    if ($Clear)
    {
        $Script:M365DSCStringReplacementMap = @{}
    }

    if ($PSBoundParameters.ContainsKey('Map'))
    {
        foreach ($key in $Map.Keys)
        {
            $Script:M365DSCStringReplacementMap[$key] = $Map[$key]
        }
    }
}

<#
.DESCRIPTION
    This function returns the string replacement map used during export.

.FUNCTIONALITY
    Internal
#>
function Get-M365DSCStringReplacementMap
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param()

    return $Script:M365DSCStringReplacementMap.Clone()
}

<#
.Description
This function generates DSC string from an exported result hashtable

.Functionality
Internal
#>
function Get-M365DSCExportContentForResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ResourceName,

        [Parameter(Mandatory = $true)]
        [System.String]
        [ValidateSet('ServicePrincipalWithThumbprint', 'ServicePrincipalWithSecret', 'ServicePrincipalWithPath', 'CredentialsWithTenantId', 'CredentialsWithApplicationId', 'Credentials', 'ManagedIdentity', 'AccessTokens')]
        $ConnectionMode,

        [Parameter(Mandatory = $true)]
        [System.String]
        $ModulePath,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Results,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String[]]
        $NoEscape,

        [Parameter()]
        [switch]
        $SkipAuthenticationUpdate,

        [Parameter()]
        [switch]
        $AllowVariablesInStrings
    )

    if (-not $SkipAuthenticationUpdate)
    {
        $withoutAuthentication = Update-M365DSCExportAuthenticationResults -ConnectionMode $ConnectionMode `
            -Results $Results
        $Results = $withoutAuthentication.Results
        $NoEscape += $withoutAuthentication.NoEscape
    }
    $NoEscape = $NoEscape | Select-Object -Unique

    $OrganizationName = ''
    if ($ConnectionMode -like 'ServicePrincipal*' -or `
            $ConnectionMode -eq 'ManagedIdentity')
    {
        $OrganizationName = $Results.TenantId
    }
    elseif ($null -ne $Credential.UserName)
    {
        $OrganizationName = $Credential.UserName.Split('@')[1]
    }
    else
    {
        $OrganizationName = ''
    }

    $primaryKey = ''
    $ModuleFullName = "MSFT_" + $ResourceName
    $Resource = $Script:AllM365DSCResources.$ResourceName
    $Keys = $Resource.Properties.Where({ $_.IsMandatory }) | `
        Select-Object -ExpandProperty Name
    if ($null -eq $keys)
    {
        if (-not (Get-Module $ModuleFullName))
        {
            Import-Module $Resource.Path -Force
        }
        $cmdInfo = Get-Command $ModuleFullName\Get-TargetResource -ErrorAction SilentlyContinue
        $Keys = $cmdInfo.Parameters.Values.Where({ $_.ParameterSets.Values.IsMandatory }).Name
    }

    if ($Keys.Contains('IsSingleInstance'))
    {
        $primaryKey = ''
    }
    elseif ($Keys.Contains('DisplayName') -and -not [System.String]::IsNullOrEmpty($Results.DisplayName))
    {
        $primaryKey = $Results.DisplayName
    }
    elseif ($Keys.Contains('Name'))
    {
        $primaryKey = $Results.Name
    }
    elseif ($Keys.Contains('Title'))
    {
        $primaryKey = $Results.Title
    }
    elseif ($Keys.Contains('Identity'))
    {
        $primaryKey = $Results.Identity
    }
    elseif ($Keys.Contains('Id'))
    {
        $primaryKey = $Results.Id
    }
    elseif ($Keys.Contains('CDNType'))
    {
        $primaryKey = $Results.CDNType
    }
    elseif ($Keys.Contains('WorkspaceName'))
    {
        $primaryKey = $Results.WorkspaceName
    }
    elseif ($Keys.Contains('OrganizationName'))
    {
        $primaryKey = $Results.OrganizationName
    }
    elseif ($Keys.Contains('DomainName'))
    {
        $primaryKey = $Results.DomainName
    }
    elseif ($Keys.Contains('UserPrincipalName'))
    {
        $primaryKey = $Results.UserPrincipalName
    }

    if ([String]::IsNullOrEmpty($primaryKey) -and `
        -not $Keys.Contains('IsSingleInstance'))
    {
        foreach ($Key in $Keys)
        {
            $primaryKey += $Results.$Key
        }
    }

    $instanceName = $ResourceName
    if (-not [System.String]::IsNullOrEmpty($primaryKey))
    {
        if ($AllowVariablesInStrings)
        {
            $primaryKey = $primaryKey.Replace('`', '``').Replace('"', '`"')
        }
        else
        {
            $primaryKey = $primaryKey.Replace('`', '``').Replace('$', '`$').Replace('"', '`"')
        }
        $primaryKey = Update-M365DSCSpecialCharacters -String $primaryKey
        $instanceName += "-$primaryKey"
    }

    if ($Results.ContainsKey('Workload'))
    {
        $instanceName += "-$($Results.Workload)"
    }

    # Check to see if a resource with this exact name was already exported, if so, append a number to the end.
    $i = 2
    $tempName = $instanceName
    while ($null -ne $Global:M365DSCExportedResourceInstancesNames -and `
           $Global:M365DSCExportedResourceInstancesNames.Contains($tempName))
    {
        $tempName = $instanceName + "-" + $i.ToString()
        $i++
    }
    $instanceName = $tempName
    [string[]]$Global:M365DSCExportedResourceInstancesNames += $tempName

    $content = [System.Text.StringBuilder]::New()
    [void]$content.Append("        $ResourceName `"$instanceName`"`r`n")
    [void]$content.Append("        {`r`n")
    $partialContent = Get-DSCBlock -Params $Results -ModulePath $ModulePath -NoEscape $NoEscape -AllowVariablesInStrings:$AllowVariablesInStrings

    if ($partialContent.ToLower().IndexOf($OrganizationName.ToLower()) -gt 0)
    {
        $partialContent = $partialContent -ireplace [regex]::Escape($OrganizationName + ':'), "`$($OrganizationName):"
        $partialContent = $partialContent -ireplace [regex]::Escape($OrganizationName), "`$OrganizationName"
        $partialContent = $partialContent -ireplace [regex]::Escape('@' + $OrganizationName), "@`$OrganizationName"
    }

    # Apply additional string to variable replacements from mapping
    if ($null -ne $Script:M365DSCStringReplacementMap -and $Script:M365DSCStringReplacementMap.Count -gt 0)
    {
        foreach ($entry in $Script:M365DSCStringReplacementMap.GetEnumerator())
        {
            $target = $entry.Key
            $varName = $entry.Value
            if ([System.String]::IsNullOrEmpty($target) -or [System.String]::IsNullOrEmpty($varName))
            {
                Write-Verbose -Message "Skipping invalid string replacement map entry: Key = '$target', VariableName = '$varName'"
                continue
            }
            # Skip if already handled as OrganizationName
            if ($OrganizationName -and ($target -ieq $OrganizationName))
            {
                Write-Verbose -Message "Skipping replacement for target [$target] because it matches the OrganizationName: '$OrganizationName'"
                continue
            }

            if ($partialContent.ToLower().IndexOf($target.ToLower()) -gt 0)
            {
                $partialContent = $partialContent -ireplace [regex]::Escape($target + ':'), "`$(`$ConfigurationData.NonNodeData.$varName):"
                $partialContent = $partialContent -ireplace [regex]::Escape($target), "`$(`$ConfigurationData.NonNodeData.$varName)"
                $partialContent = $partialContent -ireplace [regex]::Escape('@' + $target), "@`$(`$ConfigurationData.NonNodeData.$varName)"
            }
        }
    }

    [void]$content.Append($partialContent)
    [void]$content.Append("        }`r`n")

    return $content.ToString()
}

<#
.Description
This function gets all resources that support the specified authentication method and
determines the most secure authentication method supported by the resource.

.Functionality
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
            Import-Module $resource.FullName -Force
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
.Description
This function gets all available M365DSC resources in the module

.Example
Get-M365DSCAllResources

.Functionality
Public
#>
function Get-M365DSCAllResources
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    [CmdletBinding()]
    param ()

    $allResources = Get-ChildItem -Path ($PSScriptRoot + '\..\DSCResources\') -Recurse -Filter '*.psm1'
    $result = @()
    foreach ($resource in $allResources)
    {
        $result += $resource.Name -replace 'MSFT_', '' -replace '.psm1', ''
    }

    return $result
}

<#
.Description
This function checks if the specified object has the specified property

.Functionality
Internal, Hidden
#>
function Test-M365DSCObjectHasProperty
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [Object]
        $Object,

        [Parameter(Mandatory = $true, Position = 2)]
        [String]
        $PropertyName
    )

    if (([bool]($Object.PSobject.Properties.Name -contains $PropertyName)) -eq $true)
    {
        if ($null -ne $Object.$PropertyName)
        {
            return $true
        }
    }
    return $false
}

<#
.Description
    This function returns the connection workloads for the specified DSC resources

.Parameter ResourceNames
    Specifies the resources for which the connection workloads should be determined.
    Either a single string, an array of strings or an object with 'Name' and 'AuthenticationMethod' can be provided.

.Example
    Get-M365DSCConnectedWorkloadList -ResourceNames AADUser

.EXAMPLE
    Get-M365DSCConnectedWorkloadList -ResourceNames @('AADUser', 'AADGroup')

.EXAMPLE
    Get-M365DSCConnectedWorkloadList -ResourceNames @{Name = 'AADUser'; AuthenticationMethod = 'Credentials'}

.Functionality
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
    return ($workloads | Sort-Object {$_.Name})
}

<#
.Description
    This function returns the workload to which the specified DSC resources belongs.

.Parameter ResourceName
    Specifies the resources for which the workloads should be determined.
    Either a single string or an array of strings.

.Example
    Get-M365DSCWorkloadForResource -ResourceName AADUser

.EXAMPLE
    Get-M365DSCWorkloadForResource -ResourceName @('AADUser', 'AADGroup')

.Functionality
Internal
#>
function Get-M365DSCWorkloadForResource
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [System.String[]]
        $ResourceName
    )

    $workloads = @()
    foreach ($resource in $ResourceName)
    {
        foreach ($workload in $Script:M365DSCWorkloads)
        {
            if ($resource -like "$($workload)*")
            {
                if ($workloads -notcontains $workload)
                {
                    $workloads += $workload
                    break
                }
            }
        }
    }

    return $workloads | Sort-Object
}

<#
.Description
This function gets the used authentication mode based on the specified parameters

.Functionality
Internal
#>
function Get-M365DSCAuthenticationMode
{
    [CmdletBinding()]
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
.Description
This function creates Markdown documentation of all public M365DSC cmdlets
and places these in the correct location of the docs folder.

.Functionality
Internal
#>
function New-M365DSCCmdletDocumentation
{
    param()

    $cmdletDocsRoot = Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\docs\docs\user-guide\cmdlets'

    if ((Test-Path -Path $cmdletDocsRoot) -eq $false)
    {
        $null = New-Item -Path $cmdletDocsRoot -ItemType Directory
    }

    $filesInFolder = Get-ChildItem -Path $cmdletDocsRoot
    if ($filesInFolder.Count -ne 0)
    {
        Remove-Item -Path $filesInFolder.FullName -Confirm:$false
    }

    Write-Host -Object ' '
    Write-Host -Object 'Creating Markdown documentation for M365DSC cmdlets:' -ForegroundColor Gray

    $counter = 0
    foreach ($command in (Get-Module Microsoft365Dsc).ExportedCommands.GetEnumerator())
    {
        $commandName = $command.Key
        $helpInfo = Get-Help $commandName
        $functionality = $helpInfo.Functionality -split ', '
        if ('Public' -in $functionality)
        {
            Write-Host -Object "  * $commandName " -ForegroundColor Gray -NoNewline

            $output = New-Object -TypeName System.Text.StringBuilder

            $null = $output.AppendLine("# $($commandName)")
            $null = $output.AppendLine('')

            $helpInfo = Get-Help -Name $commandName
            if ($helpInfo.description.Count -ne 0)
            {
                $null = $output.AppendLine('## Description')
                $null = $output.AppendLine('')
                $null = $output.AppendLine($helpInfo.Description[0].Text)
                $null = $output.AppendLine('')
            }

            $cmd = Get-Command -Name $commandName
            if ([String]::IsNullOrEmpty($cmd.OutputType) -eq $false)
            {
                $null = $output.AppendLine('## Output')
                $null = $output.AppendLine('')
                $null = $output.AppendLine('This function outputs information as the following type:')
                $null = $output.AppendLine("**$($cmd.OutputType)**")
                $null = $output.AppendLine('')
            }
            else
            {
                $null = $output.AppendLine('## Output')
                $null = $output.AppendLine('')
                $null = $output.AppendLine('This function does not generate any output.')
                $null = $output.AppendLine('')
            }

            $ast = $cmd.ScriptBlock.Ast
            $parameters = $null
            $parameters = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)

            $null = $output.AppendLine('## Parameters')
            $null = $output.AppendLine('')
            if ($parameters.Count -gt 0)
            {
                $null = $output.AppendLine('| Parameter | Required | DataType | Default Value | Allowed Values | Description |')
                $null = $output.AppendLine('| --- | --- | --- | --- | --- | --- |')

                $ast = $cmd.ScriptBlock.Ast
                $parameters = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
                foreach ($parameter in $parameters)
                {
                    $paramName = $parameter.Name.VariablePath.UserPath

                    $paramHelp = $helpInfo.parameters.parameter | Where-Object { $_.Name -eq $paramName }
                    $description = ''
                    if ($paramHelp.description.Count -gt 0)
                    {
                        $description = $paramHelp.description[0].Text
                    }
                    $mandatory = $parameter.Attributes.Where({ $_.TypeName.FullName -eq 'Parameter' }).NamedArguments.Where({ $_.ArgumentName -eq 'Mandatory' }).Argument.VariablePath.UserPath
                    if ($null -eq $mandatory)
                    {
                        $mandatory = 'False'
                    }
                    $mandatory = (Get-Culture).TextInfo.ToTitleCase($mandatory.ToLower())

                    $null = $output.AppendLine("| $($paramName) | $($mandatory) | $($parameter.StaticType.Name) | $($parameter.DefaultValue.Value) | $($parameter.Attributes.Where({$_.TypeName.FullName -eq 'ValidateSet'}).PositionalArguments.Value -join ', ') | $($description) |")

                }
                $null = $output.AppendLine('')
            }
            else
            {
                $null = $output.AppendLine('This function does not have any input parameters.')
            }

            if ($helpInfo.examples.example.Count -ne 0)
            {
                $null = $output.AppendLine('## Examples')
                $null = $output.AppendLine('')
                foreach ($example in $helpInfo.examples.example)
                {
                    $null = $output.AppendLine($example.title)
                    $null = $output.AppendLine('')
                    $null = $output.AppendLine("``$($example.code)``")
                    $null = $output.AppendLine('')
                }
            }

            $savePath = Join-Path -Path $cmdletDocsRoot -ChildPath "$commandName.md"
            $null = Out-File `
                -InputObject ($output.ToString() -replace '\r?\n', "`r`n") `
                -FilePath $savePath `
                -Encoding utf8 `
                -Force:$Force
            Write-Host -Object $Global:M365DSCEmojiGreenCheckmark -ForegroundColor Gray
            $counter++
        }
    }

    Write-Host -Object ' '
    Write-Host -Object "Total number files created: $counter" -ForegroundColor Gray
    Write-Host -Object ' '
}

<#
.Description
This function creates an example from the resource schema, using ReverseDSC code.

.Parameter ResourceName
Specifies the resource name for which the example should be generated.

.Functionality
Internal, Hidden
#>
function New-M365DSCResourceExample
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ResourceName
    )

    $resource = Get-DscResourceV2 -Name $ResourceName
    $params = Get-DSCFakeParameters -ModulePath $resource.Path
    $params.Credential = '$Credscredential'

    if ($params.ContainsKey('ApplicationId'))
    {
        $params.Remove('ApplicationId')
    }

    if ($params.ContainsKey('TenantId'))
    {
        $params.Remove('TenantId')
    }

    if ($params.ContainsKey('ApplicationSecret'))
    {
        $params.Remove('ApplicationSecret')
    }

    if ($params.ContainsKey('CertificateThumbprint'))
    {
        $params.Remove('CertificateThumbprint')
    }

    if ($params.ContainsKey('CertificatePath'))
    {
        $params.Remove('CertificatePath')
    }

    if ($params.ContainsKey('CertificatePassword'))
    {
        $params.Remove('CertificatePassword')
    }

    [string]$userName = 'admin@contoso.onmicrosoft.com'
    [string]$userPassword = 'dummypassword'
    [securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
    [pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)

    $resourceExample = Get-M365DSCExportContentForResource -ResourceName $ResourceName -ModulePath $resource.Path -Results $params -ConnectionMode Credentials -Credential $credObject

    $resourceExample = $resourceExample.TrimEnd() -replace ';', ''

    $exampleText = @"
<#
This example is used to test new resources and showcase the usage of new resources being worked on.
It is not meant to use as a production baseline.
#>

Configuration Example
{
    param
    (
        [Parameter(Mandatory = `$true)]
        [PSCredential]
        `$Credscredential
    )
    Import-DscResource -ModuleName Microsoft365DSC

    node localhost
    {
$resourceExample
    }
}
"@

    return $exampleText
}

<#
.Description
This function creates an example from the resource schema, using ReverseDSC code.

.Parameter ResourceName
Specifies the resource name for which the example should be generated.

.Functionality
Internal
#>
function New-M365DSCMissingResourcesExample
{
    $location = $PSScriptRoot

    $m365Resources = Get-DscResourceV2 -Module 'Microsoft365DSC' | Select-Object -ExpandProperty Name
    $examplesPath = Join-Path $location -ChildPath '..\..\..\Examples\Resources'
    $examples = Get-ChildItem -Path $examplesPath | Where-Object { $_.PsIsContainer } | Select-Object -ExpandProperty Name

    [array]$differences = Compare-Object -ReferenceObject $m365Resources -DifferenceObject $examples

    $count = 1
    $total = $differences.Count

    foreach ($difference in $differences)
    {
        Write-Host "[$count/$total] Processing $($difference.InputObject)"
        $path = Join-Path -Path '.\Examples\Resources' -ChildPath $difference.InputObject
        switch ($difference.SideIndicator)
        {
            '<='
            {
                Write-Host '  - Example missing, generating!'
                $null = New-Item -Path $path -ItemType Directory
                $exampleFile = Join-Path -Path $path -ChildPath '1-Configure.ps1'
                Set-Content -Path $exampleFile -Value (New-M365DSCResourceExample -ResourceName $difference.InputObject)
            }
            '=>'
            {
                Write-Host '  - No resource for existing example, removing!'
                Remove-Item -Path $path -Force -Confirm:$false
            }
        }
        $count++
    }
}

<#
.Description
This function removes the authentication parameters from the hashtable.

.Functionality
Internal
#>
function Remove-M365DSCAuthenticationParameter
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $BoundParameters
    )

    if ($BoundParameters.ContainsKey('Ensure'))
    {
        $BoundParameters.Remove('Ensure') | Out-Null
    }
    if ($BoundParameters.ContainsKey('Credential'))
    {
        $BoundParameters.Remove('Credential') | Out-Null
    }
    if ($BoundParameters.ContainsKey('ApplicationId'))
    {
        $BoundParameters.Remove('ApplicationId') | Out-Null
    }
    if ($BoundParameters.ContainsKey('ApplicationSecret'))
    {
        $BoundParameters.Remove('ApplicationSecret') | Out-Null
    }
    if ($BoundParameters.ContainsKey('TenantId'))
    {
        $BoundParameters.Remove('TenantId') | Out-Null
    }
    if ($BoundParameters.ContainsKey('CertificatePassword'))
    {
        $BoundParameters.Remove('CertificatePassword') | Out-Null
    }
    if ($BoundParameters.ContainsKey('CertificatePath'))
    {
        $BoundParameters.Remove('CertificatePath') | Out-Null
    }
    if ($BoundParameters.ContainsKey('CertificateThumbprint'))
    {
        $BoundParameters.Remove('CertificateThumbprint') | Out-Null
    }
    if ($BoundParameters.ContainsKey('ManagedIdentity'))
    {
        $BoundParameters.Remove('ManagedIdentity') | Out-Null
    }
    if ($BoundParameters.ContainsKey('Verbose'))
    {
        $BoundParameters.Remove('Verbose') | Out-Null
    }
    if ($BoundParameters.ContainsKey('AccessTokens'))
    {
        $BoundParameters.Remove('AccessTokens') | Out-Null
    }
    return $BoundParameters
}

<#
.Description
This function analyzes an M365DSC configuration file and returns information about potential issues (e.g., duplicate primary keys).

.Example
Get-M365DSCConfigurationConflict -ConfigurationContent "content"

.Functionality
Public
#>
function Get-M365DSCConfigurationConflict
{
    [CmdletBinding()]
    [OutputType([Array])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ConfigurationContent
    )

    $results = @()
    Write-Verbose -Message "Converting configuration's content into a PowerShell Object using DSCParser"
    $parsedContent = ConvertTo-DSCObject -Content $ConfigurationContent

    $resourcesPrimaryIdentities = @()
    $resourcesInModule = Get-DscResourceV2 -Module 'Microsoft365DSC'
    foreach ($component in $parsedContent)
    {
        $resourceDefinition = $resourcesInModule | Where-Object -FilterScript {$_.Name -eq $component.ResourceName}
        [Array]$mandatoryProperties = $resourceDefinition.Properties | Where-Object -FilterScript {$_.IsMandatory}
        $primaryKeyValues = ""
        foreach ($mandatoryKey in $mandatoryProperties.Name)
        {
            $primaryKeyValues += "$($component.$mandatoryKey)|"
        }
        $entryValue = "[$($component.ResourceName)]$primaryKeyValues"
        if ($resourcesPrimaryIdentities.Contains($entryValue))
        {
            Write-Verbose -Message "Found primary key conflict in resource {$($component.ResourceInstanceName)}"
            $currentEntry = @{
                ResourceName         = $component.ResourceName
                InstanceName         = $component.ResourceInstanceName
                AdditionalProperties = @{}
                Reason               = "DuplicatePrimaryKey"
            }

            foreach ($mandatoryKey in $mandatoryProperties.Name)
            {
                $currentEntry.AdditionalProperties.Add($mandatoryKey, $component.$mandatoryKey)
            }
            $results += $currentEntry
        }
        else
        {
            $resourcesPrimaryIdentities += $entryValue
        }
    }
    return $results
}

<#
.DESCRIPTION
    This function returns a hashtable with aligned to the parameter pattern of the given cmdlet.

.EXAMPLE
    PS> $param = @{
            Path = 'C:\Test'
            DoesNotExist = '123'
        }
    PS> Sync-M365DSCParameter -Command (Get-Command -Name Get-ChildItem) -Parameters $param

.FUNCTIONALITY
    Private
#>
# TODO: Check if necessary to keep this function.
function Sync-M365DSCParameter
{
    [Cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript( {
                    $_ -is [System.Management.Automation.FunctionInfo] -or
                    $_ -is [System.Management.Automation.CmdletInfo] -or
                    $_ -is [System.Management.Automation.ExternalScriptInfo] -or
                    $_ -is [System.Management.Automation.AliasInfo]
            })]
        [object]$Command,

        [hashtable]$Parameters,

        [switch]$ConvertValue
    )

    if (-not $PSBoundParameters.ContainsKey('Parameters'))
    {
        $Parameters = ([hashtable]$ALBoundParameters).Clone()
    }
    else
    {
        $Parameters = ([hashtable]$Parameters).Clone()
    }

    if ($Command -is [System.Management.Automation.AliasInfo])
    {
        $command = & (Get-Module -Name $Command.Source) {
            param([string]$Name)

            Get-Command -Name $Name
        } $Command.Definition
    }

    $commonParameters = [System.Management.Automation.Internal.CommonParameters].GetProperties().Name
    $commandParameterKeys = $Command.Parameters.Keys.GetEnumerator() | ForEach-Object { $_ }
    $parameterKeys = $Parameters.Keys.GetEnumerator() | ForEach-Object { $_ }

    $keysToRemove = Compare-Object -ReferenceObject $commandParameterKeys -DifferenceObject $parameterKeys |
        Select-Object -ExpandProperty InputObject

    $keysToRemove = $keysToRemove + $commonParameters | Select-Object -Unique #remove the common parameters

    foreach ($key in $keysToRemove)
    {
        $Parameters.Remove($key)
    }

    if ($ConvertValue.IsPresent)
    {
        $keysToUpdate = @{}
        foreach ($kvp in $Parameters.GetEnumerator())
        {
            if (-not $kvp.Value) # $null or empty string will not trip up conversion
            {
                continue
            }

            $targetType = $Command.Parameters[$kvp.Key].ParameterType
            $sourceType = $kvp.Value.GetType()
            $targetValue = $kvp.Value -as $targetType

            if (-not $targetValue -and $targetType.ImplementedInterfaces -contains [Collections.IList])
            {
                $targetValue = $targetType::new()
                foreach ($v in $kvp.Value)
                {
                    $targetValue.Add($v)
                }
            }

            if (-not $targetValue -and $targetType.ImplementedInterfaces -contains [Collections.IDictionary] )
            {
                $targetValue = $targetType::new()
                foreach ($k in $kvp.Value.GetEnumerator())
                {
                    $targetValue.Add($k.Key, $k.Value)
                }
            }

            if (-not $targetValue -and ($sourceType.ImplementedInterfaces -contains [Collections.IList] -and $targetType.ImplementedInterfaces -notcontains [Collections.IList]))
            {
                Write-Verbose -Message "Value of source parameter $($kvp.Key) is a collection, but target parameter is not. Selecting first object"
                $targetValue = $kvp.Value | Select-Object -First 1
            }

            if (-not $targetValue)
            {
                Write-Error -Message "Conversion of source parameter $($kvp.Key) (Type: $($sourceType.FullName)) to type $($targetType.FullName) was impossible"
                return
            }

            $keysToUpdate[$kvp.Key] = $targetValue
        }
    }

    if ($keysToUpdate)
    {
        foreach ($kvp in $keysToUpdate.GetEnumerator())
        {
            $Parameters[$kvp.Key] = $kvp.Value
        }
    }

    if ($PSBoundParameters.ContainsKey('Parameters'))
    {
        $Parameters
    }
}

<#
.SYNOPSIS
    Joins two or more M365DSC configurations into a single configuration.

.DESCRIPTION
    This function is used to join two or more M365DSC configurations into a single configuration.
    The function reads the configuration from the specified paths and combines them into a single configuration.
    Please note that the function won't be updating the authentication parameters if they differ between the configurations. Make sure that the authentication parameters are the same over all configurations.

.PARAMETER ConfigurationFile
    The name of the first configuration file to use as the base configuration.

.PARAMETER ConfigurationPath
    The directory path to the configuration files to join to the base configuration.

.EXAMPLE
    Join-M365DSCConfiguration -ConfigurationFile 'M365TenantConfig.ps1' -ConfigurationPath 'D:\testbed'
    This example joins the 'M365TenantConfig.ps1' file with all the configuration files in the 'D:\testbed' directory.

.FUNCTIONALITY
    Public
#>
function Join-M365DSCConfiguration
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $ConfigurationFile,

        [Parameter(Mandatory = $true)]
        [string]
        $ConfigurationPath
    )

    if ($ConfigurationFile -notlike "*.ps1")
    {
        throw "The ConfigurationFile parameter must be a .ps1 file."
    }

    if (-not (Test-Path -Path $ConfigurationPath))
    {
        throw "The ConfigurationPath parameter must be a valid path."
    }

    $ConfigurationFilePath = Join-Path -Path $ConfigurationPath -ChildPath $ConfigurationFile
    $ConfigurationPath = Join-Path -Path $ConfigurationPath -ChildPath "*"

    $baseConfiguration = ConvertTo-DSCObject -Path $ConfigurationFilePath
    $additionalConfigurations = Get-Item -Path $ConfigurationPath -Filter *.ps1 -Exclude $ConfigurationFile | ForEach-Object { ConvertTo-DSCObject -Path $_.FullName }

    $combinedArray = @($baseConfiguration) + @($additionalConfigurations)
    $combinedConfiguration = ConvertFrom-DSCObject -DSCResources $combinedArray

    # Indent all lines by 8 spaces to match the indentation of the configuration file
    $combinedConfiguration = $combinedConfiguration -replace '(?m)^', '        '
    $combinedConfiguration = $combinedConfiguration.TrimEnd()

    # Remove everything in the "Node localhost" part in the configuration file, while excluding the last two closing brackets
    $content = Get-Content -Path $ConfigurationFilePath -Raw
    $content = $content -replace '(?s)(?<=Node localhost\s*\{)(.*\s{8}\}?)(?=\s*\})', ''

    # Append the combined configuration after the "Node localhost" part in the configuration file
    $content = $content -replace '(?s)(?<=Node localhost\s*\{)', "`r`n$combinedConfiguration"

    return $content
}

<#
.DESCRIPTION
    Invokes a script-based DSC resource from a Windows PowerShell 5.1 session into a PowerShell Core session.

.PARAMETER Path
    The path to the module containing the resource.

.PARAMETER FunctionName
    The name of the function to invoke.

.PARAMETER Parameters
    The parameters to pass to the function.

.EXAMPLE
    Invoke-PowerShellCoreResource -Path 'C:\Program Files\...\DSCResources\MSFT_Resource\MSFT_Resource.psm1' -FunctionName Test-TargetResource -Parameters @{ Name = 'Value' }

.EXAMPLE
    # From inside of a DSC resource
    Invoke-PowerShellCoreResource -Path $PSCommandPath -FunctionName $MyInvocation.MyCommand.Name -Parameters $PSBoundParameters

.FUNCTIONALITY
    Internal

.OUTPUTS
    Result of the invoked function.
#>
function Invoke-PowerShellCoreResource
{
    [CmdletBinding()]
    [OutputType([System.Object])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Path', Justification = 'Using statement not detected')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'FunctionName', Justification = 'Using statement not detected')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Parameters', Justification = 'Using statement not detected')]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Get-TargetResource', 'Set-TargetResource', 'Test-TargetResource', 'Export-TargetResource')]
        [System.String]$FunctionName,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]$Parameters
    )

    if (-not $script:PSCoreSessionInitialized)
    {
        Initialize-PowerShellCoreSession
    }

    $output = Invoke-Command -Session $PSCoreSession -ScriptBlock {
        Import-Module -Name $using:Path
        & $using:FunctionName @using:Parameters
    }

    return $output
}

<#
.DESCRIPTION
    Initializes a PowerShell Core session for use with Invoke-PowerShellCoreResource.

.FUNCTIONALITY
    Private

.EXAMPLE
    Initialize-PowerShellCoreSession
#>
function Initialize-PowerShellCoreSession
{
    [CmdletBinding()]
    param ()

    $script:PSCoreSession = New-PSSession -ComputerName localhost -ConfigurationName PowerShell.7 -EnableNetworkAccess
    $lcmConfig = Get-DscLocalConfigurationManager
    Invoke-Command -Session $script:PSCoreSession -ScriptBlock {
        Import-Module -Name Microsoft365DSC -Alias @() -Cmdlet @() -Variable @() -DisableNameChecking -SkipEditionCheck
        Set-M365DSCLCMConfiguration -LCMConfig $using:lcmConfig
    }
    $script:PSCoreSessionInitialized = $true
}

<#
.Description
This function writes messages to the console or verbose output.

.PARAMETER Message
Specifies the message to write.

.PARAMETER DeferWrite
Specifies if writing the message should be deferred. Adheres to -NoNewLine behavior of Write-Host.

.PARAMETER CommitWrite
Specifies if cached messages of -DeferWrite should be combined and written.
Combining of the messages is done by joining them without any characters between.

.EXAMPLE
Write-M365DSCHost -Message "This is a message."

.Functionality
Internal
#>
function Write-M365DSCHost
{
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param
    (
        [Parameter(Position = 0)]
        [System.String]
        $Message,

        [Parameter()]
        [ConsoleColor]
        $ForegroundColor = [System.Console]::ForegroundColor,

        [Parameter(ParameterSetName = 'DeferWrite')]
        [switch]
        $DeferWrite,

        [Parameter(ParameterSetName = 'CommitWrite')]
        [switch]
        $CommitWrite
    )

    if (-not [System.String]::IsNullOrEmpty($Message))
    {
        if ($null -eq $Script:M365DSCHostMessages)
        {
            $Script:M365DSCHostMessages = @()
        }

        if ($DeferWrite)
        {
            $Script:M365DSCHostMessages += @{
                Message = $Message
                ForegroundColor = $ForegroundColor
            }
            return
        }

        if ([Environment]::UserInteractive)
        {
            if ($CommitWrite -and $Script:M365DSCHostMessages.Count -gt 0)
            {
                for ($i = 0; $i -lt $Script:M365DSCHostMessages.Count - 1; $i++)
                {
                    Write-Host -Object $Script:M365DSCHostMessages[$i].Message -ForegroundColor $Script:M365DSCHostMessages[$i].ForegroundColor -NoNewline
                }
                Write-Host -Object $Script:M365DSCHostMessages[-1].Message -ForegroundColor $Script:M365DSCHostMessages[-1].ForegroundColor -NoNewline
                $Script:M365DSCHostMessages = @()
            }

            if (-not [System.String]::IsNullOrEmpty($Message))
            {
                Write-Host -Object $Message -ForegroundColor $ForegroundColor
            }
        }
        else
        {
            $outputMessage = ''
            if ($CommitWrite)
            {
                $outputMessage += $Script:M365DSCHostMessages.Message -join ''
                $Script:M365DSCHostMessages = @()
            }
            $finalMessage = $outputMessage + $Message
            if (-not [System.String]::IsNullOrEmpty($Message))
            {
                Write-Verbose -Message $finalMessage -Verbose
            }
        }
    }
}

<#
.DESCRIPTION
    This function sends a batch request to the Microsoft Graph API.

.PARAMETER Requests
    An array of hashtables representing the requests to be sent in the batch.
    A request hashtable should contain the following keys:
    - id: A unique identifier for the request.
    - method: The HTTP method to use (e.g., GET, POST).
    - url: The API endpoint URL.

.EXAMPLE
    $requests = @(
        @{
            id = '1'
            method = 'GET'
            url = '/users'
        }
    )
    Invoke-M365DSCGraphBatchRequest -Requests $requests
#>
function Invoke-M365DSCGraphBatchRequest
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable[]])]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Hashtable[]]
        $Requests,

        [Parameter()]
        [switch]
        $AsList
    )

    $batchResponses = [System.Collections.Generic.List[System.Collections.Hashtable]]::new()
    for ($i = 0; $i -lt $Requests.Count; $i += 20)
    {
        $batchRequestSized = $Requests[$i..([Math]::Min($i + 19, $Requests.Count - 1))]

        $request = @{
            requests = $batchRequestSized
        }

        Write-Verbose -Message "Sending BATCH Request with:`r`n$($request | ConvertTo-Json -Depth 10))"
        $batchResponses.AddRange([System.Collections.Hashtable[]](Invoke-MgGraphRequest -Method POST `
            -Uri 'beta/$batch' `
            -Body ($request | ConvertTo-Json -Depth 10) `
            -ErrorAction SilentlyContinue).responses)
    }

    if ($AsList)
    {
        return $batchResponses
    }
    return $batchResponses.ToArray()
}

<#
.DESCRIPTION
    This function splits a large M365DSC configuration file into smaller files based on size and resource count limits.

.PARAMETER Path
    The path to the M365DSC configuration file to split.

.PARAMETER OutputFolder
    The folder where the split configuration files will be saved. Defaults to the same folder as the input file.

.PARAMETER MaxFileSizeMB
    The maximum size (in megabytes) for each split configuration file. Default is 3 MB.

.PARAMETER MaxResources
    The maximum number of resources per split configuration file. Default is 0 (no limit).

.EXAMPLE
    Split-M365DSCConfiguration -Path 'C:\Configs\M365TenantConfig.ps1' -OutputFolder 'C:\Configs\Split' -MaxFileSizeMB 2 -MaxResources 50
    This example splits the 'M365TenantConfig.ps1' file into smaller files, each with a maximum size of 2 MB and a maximum of 50 resources, saving them in the 'C:\Configs\Split' folder.

.FUNCTIONALITY
    Public
#>
function Split-M365DSCConfiguration {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Path,

        [Parameter()]
        [System.String]
        $OutputFolder = (Split-Path $Path),

        [Parameter()]
        [System.Double]
        $MaxFileSizeMB = 3,

        [Parameter()]
        [System.Int32]
        $MaxResources = 0  # 0 = ignore resource count limit
    )

    $fileContent = Get-Content -Encoding utf8 -Path $Path -Raw

    # Extract content inside "Node localhost { ... }"
    $pattern = 'Node localhost\s*{([\s\S]*)\s+}(\r|\n)+\s+}'
    $nodeMatch = [regex]::Match($fileContent, $pattern)
    if (-not $nodeMatch.Success) {
        throw "Could not find a 'Node localhost { ... }' block in file: $Path"
    }

    $nodeContent = $nodeMatch.Groups[1].Value

    # Extract header (everything before Node localhost)
    $header = ($fileContent -split 'Node localhost')[0] + "Node localhost`n    {`n"
    $footer = "`n    }`n}`n`nM365TenantConfig -ConfigurationData .\ConfigurationData.psd1"

    # Split into DSC resource text blocks using brace-depth parsing
    $resources = @()
    $lines = $nodeContent -split "`r?`n"
    $currentResource = [System.Text.StringBuilder]::new()
    $braceDepth = 0
    $insideResource = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        # Detect resource start
        if (-not $insideResource -and $line.Trim() -match '^[a-zA-Z0-9_]+\s+"[^"]+"') {
            $insideResource = $true
            $null = $currentResource.Clear()
            $null = $currentResource.AppendLine($line)
            # Calculate brace depth
            $braceDepth = ($line -split '{').Count - ($line -split '}').Count
            continue
        }

        if ($insideResource) {
            $null = $currentResource.AppendLine($line)

            # Adjust brace depth based on line content
            $braceDepth += ($line -split '{').Count - ($line -split '}').Count

            # End of resource block
            if ($braceDepth -le 0) {
                $resources += "        " + $currentResource.ToString().Trim()
                $insideResource = $false
            }
        }
    }

    if (-not $resources) {
        throw "No DSC resources found in the Node block."
    }

    # Splitting logic
    $i = 1
    $currentGroup = @()
    $currentSize = 0
    $maxBytes = $MaxFileSizeMB * 1MB

    foreach ($res in $resources) {
        # Calculate size of the resource in bytes
        $resBytes = [System.Text.Encoding]::UTF8.GetByteCount($res)
        $resourceCountLimitReached = ($MaxResources -gt 0 -and $currentGroup.Count -ge $MaxResources)
        $sizeLimitReached = ($currentSize + $resBytes) -gt $maxBytes

        # Write current group if limits are reached
        if (($sizeLimitReached -or $resourceCountLimitReached) -and $currentGroup.Count -gt 0) {
            $outPath = Join-Path $OutputFolder ("M365TenantConfig_{0}.ps1" -f $i)
            $configText = $header + ($currentGroup -join "`n") + $footer
            Set-Content -Path $outPath -Value $configText -Encoding UTF8 -Force
            Write-M365DSCHost -Message "Created: $outPath" -CommitWrite
            $i++
            $currentGroup = @()
            $currentSize = 0
        }

        $currentGroup += $res
        $currentSize += $resBytes
    }

    # Write final group
    if ($currentGroup.Count -gt 0) {
        $outPath = Join-Path $OutputFolder ("M365TenantConfig_{0}.ps1" -f $i)
        $configText = $header + ($currentGroup -join "`n`n") + $footer
        Set-Content -Path $outPath -Value $configText -Encoding UTF8 -Force
        Write-M365DSCHost -Message "Created: $outPath" -CommitWrite
    }
}

<#
.Description
    This function retrieves the comparison metadata for a given M365DSC resource.
    The metadata indicates whether a resource requires custom comparison logic and
    should expose a Get-CompareParameters function.

.Parameter ResourceName
    The name of the M365DSC resource (without MSFT_ prefix).

.Example
    PS> Get-M365DSCResourceComparisonMetadata -ResourceName 'AADRoleAssignmentScheduleRequest'

.Functionality
    Internal
#>
function Get-M365DSCResourceComparisonMetadata
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ResourceName
    )

    if ($null -eq $Script:M365DSCComparisonMetadata)
    {
        $metadataPath = Join-Path -Path $PSScriptRoot -ChildPath '..\ComparisonMetadata.json'
        if (Test-Path -Path $metadataPath)
        {
            try
            {
                $metadataContent = Get-Content -Path $metadataPath -Raw | ConvertFrom-Json
                $Script:M365DSCComparisonMetadata = @{}
                foreach ($resource in $metadataContent.Resources.PSObject.Properties)
                {
                    $Script:M365DSCComparisonMetadata[$resource.Name] = @{
                        HasCustomComparison = $resource.Value.HasCustomComparison
                        Description = $resource.Value.Description
                    }
                }
            }
            catch
            {
                Write-Warning -Message "Failed to load comparison metadata from $metadataPath : $_"
                $Script:M365DSCComparisonMetadata = @{}
            }
        }
        else
        {
            Write-Verbose -Message "Comparison metadata file not found at $metadataPath"
            $Script:M365DSCComparisonMetadata = @{}
        }
    }

    if ($Script:M365DSCComparisonMetadata.ContainsKey($ResourceName))
    {
        return $Script:M365DSCComparisonMetadata[$ResourceName]
    }

    return @{
        HasCustomComparison = $false
    }
}

<#
.Description
    This function retrieves the comparison parameters from a resource's Get-CompareParameters function.
    This is used during drift detection and reporting to ensure that resource-specific comparison logic
    (such as PostProcessing scripts and ExcludedProperties) is applied consistently.

.Parameter ResourceName
    The name of the M365DSC resource (without MSFT_ prefix).

.Example
    PS> Get-M365DSCResourceComparisonParameters -ResourceName 'AADRoleAssignmentScheduleRequest'

.Functionality
    Internal
#>
function Get-M365DSCResourceComparisonParameters
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ResourceName
    )

    $compareParameters = @{}

    try
    {
        # Check if this resource has custom comparison logic
        $metadata = Get-M365DSCResourceComparisonMetadata -ResourceName $ResourceName

        if (-not $metadata.HasCustomComparison)
        {
            Write-Verbose -Message "Resource $ResourceName does not have custom comparison logic."
            return $compareParameters
        }

        # Import the resource module if not already loaded
        $moduleName = "MSFT_$ResourceName"
        $module = Get-Module -Name $moduleName
        $moduleConfig = Get-M365DSCModuleConfiguration

        if ($null -eq $module)
        {
            $resourceModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\DSCResources\$moduleName\$moduleName.psm1"
            if (Test-Path -Path $resourceModulePath)
            {
                $previousValue = $moduleConfig.skipModuleDependencyValidation
                if (-not $metadata.RequiresModuleCheck)
                {
                    Set-M365DSCModuleConfiguration -Key "skipModuleDependencyValidation" -Value $true
                }
                Import-Module -Name $resourceModulePath -Force -Global -Function Get-CompareParameters -Alias @() -Cmdlet @() -Variable @() -DisableNameChecking
                Set-M365DSCModuleConfiguration -Key "skipModuleDependencyValidation" -Value $previousValue
                Write-Verbose -Message "Imported module $moduleName from $resourceModulePath"
            }
            else
            {
                Write-Warning -Message "Resource module not found at $resourceModulePath"
                return $compareParameters
            }
        }

        if ($null -eq $Script:CompareParametersCache)
        {
            $Script:CompareParametersCache = @{}
        }

        if ($Script:CompareParametersCache.ContainsKey($ResourceName))
        {
            return $Script:CompareParametersCache[$ResourceName]
        }

        # Check if the Get-CompareParameters function exists
        $getCompareParamsCommand = Get-Command -Name "$moduleName\Get-CompareParameters" -ErrorAction SilentlyContinue

        if ($null -eq $getCompareParamsCommand)
        {
            Write-Warning -Message "Resource $ResourceName is marked as having custom comparison, but Get-CompareParameters function not found."
            return $compareParameters
        }

        # Invoke the Get-CompareParameters function
        $compareParameters = & "$moduleName\Get-CompareParameters"

        # Cache the retrieved parameters
        $Script:CompareParametersCache[$ResourceName] = $compareParameters
    }
    catch
    {
        Write-Warning -Message "Failed to retrieve comparison parameters for $ResourceName : $_"
    }

    return $compareParameters
}

Export-ModuleMember -Function @(
    'Assert-M365DSCBlueprint',
    'Confirm-ImportedCmdletIsAvailable',
    'Convert-M365DscHashtableToString',
    'ConvertTo-SPOUserProfilePropertyInstanceString',
    'Export-M365DSCConfiguration',
    'Get-AllSPOPackages',
    'Get-M365DSCAllResources',
    'Get-M365DSCAllResourcesDictionary',
    'Get-M365DSCAPIEndpoint',
    'Get-M365DSCArrayFromProperty',
    'Get-M365DSCAuthenticationMode',
    'Get-M365DSCComponentsWithMostSecureAuthenticationType',
    'Get-M365DSCConfigurationConflict',
    'Get-M365DSCConnectedWorkloadList',
    'Get-M365DSCExportContentForResource',
    'Get-M365DSCOrganization',
    'Get-M365DSCResourceComparisonMetadata',
    'Get-M365DSCResourceComparisonParameters',
    'Get-M365DSCResourcesByExportMode',
    'Get-M365DSCStringReplacementMap',
    'Get-M365DSCTelemetryConnectionParameter',
    'Get-M365DSCTenantDomain',
    'Get-M365DSCTenantNameFromParameterSet',
    'Get-M365DSCWorkloadForResource',
    'Get-M365TenantName',
    'Get-SPOAdministrationUrl',
    'Get-TeamByName',
    'Initialize-M365DSCAllResourcesDictionary',
    'Install-M365DSCDevBranch',
    'Invoke-M365DSCGraphBatchRequest',
    'Invoke-PowerShellCoreResource',
    'Join-M365DSCConfiguration',
    'New-M365DSCCmdletDocumentation',
    'New-M365DSCConnection',
    'New-M365DSCMissingResourcesExample',
    'Remove-M365DSCAuthenticationParameter',
    'Remove-NullEntriesFromHashtable',
    'Set-M365DSCAllResourcesDictionary',
    'Set-M365DSCStringReplacementMap',
    'Split-M365DSCConfiguration',
    'Sync-M365DSCParameter',
    'Test-CodePage',
    'Test-M365DSCParameterState',
    'Test-M365DSCTargetResource',
    'Update-M365DSCExportAuthenticationResults',
    'Write-M365DSCHost'
)

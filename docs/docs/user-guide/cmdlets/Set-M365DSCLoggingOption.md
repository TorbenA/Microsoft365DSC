# Set-M365DSCLoggingOption

## Description

This function configures the option for logging events into the Event Log.

## Output

This function does not generate any output.

## Parameters

| Parameter | Required | DataType | Default Value | Allowed Values | Description |
| --- | --- | --- | --- | --- | --- |
| IncludeNonDrifted | False | Boolean |  |  | Determines whether or not we should log information about resource's instances that don't have drifts. |

## Examples

-------------------------- EXAMPLE 1 --------------------------

`Set-M365DSCLoggingOption`

-------------------------- EXAMPLE 2 --------------------------

`Set-M365DSCLoggingOption -IncludeNonDrifted $true`

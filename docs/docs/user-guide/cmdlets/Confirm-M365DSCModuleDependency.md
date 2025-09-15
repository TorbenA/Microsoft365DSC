# Confirm-M365DSCModuleDependency

## Description

This function checks the required dependencies for a specific M365DSC module and validates that they are loaded.

## Output

This function does not generate any output.

## Parameters

| Parameter | Required | DataType | Default Value | Allowed Values | Description |
| --- | --- | --- | --- | --- | --- |
| ModuleName | True | String |  |  | The name of the DSC resource for which to check dependencies. |

## Examples

-------------------------- EXAMPLE 1 --------------------------

`Confirm-M365DSCModuleDependency -ModuleName 'MSFT_AADApplication'`



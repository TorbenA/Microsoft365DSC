# Get-M365DSCConnectedWorkloadList

## Description

This function returns the connection workloads for the specified DSC resources

## Output

This function outputs information as the following type:
**System.Collections.Hashtable[]**

## Parameters

| Parameter | Required | DataType | Default Value | Allowed Values | Description |
| --- | --- | --- | --- | --- | --- |
| ResourceNames | True | Array |  |  | Specifies the resources for which the connection workloads should be determined.
Either a single string, an array of strings or an object with 'Name' and 'AuthenticationMethod' can be provided. |

## Examples

-------------------------- EXAMPLE 1 --------------------------

`Get-M365DSCConnectedWorkloadList -ResourceNames AADUser`

-------------------------- EXAMPLE 2 --------------------------

`Get-M365DSCConnectedWorkloadList -ResourceNames @('AADUser', 'AADGroup')`

-------------------------- EXAMPLE 3 --------------------------

`Get-M365DSCConnectedWorkloadList -ResourceNames @{Name = 'AADUser'; AuthenticationMethod = 'Credentials'}`

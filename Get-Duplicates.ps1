<#
    .SYNOPSIS
        Gets duplicates or unique values in a collection.

    .DESCRIPTION
        The Get-Duplicates.ps1 script takes a collection and returns
        the duplicates (by default) or unique members (use the Unique
        switch parameter).

    .PARAMETER  Items
        Enter a collection of items. You can also pipe the items to
        Get-Duplicates.ps1.

    .PARAMETER  Unique
        Returns unique items instead of duplicates. By default, Get-Duplicates.ps1
        returns only duplicates.

    .EXAMPLE
        PS C:\> .\Get-Duplicates.ps1 -Items 1,2,3,2,4
        2

    .EXAMPLE
        PS C:\> 1,2,3,2,4 | .\Get-Duplicates.ps1
        2

    .EXAMPLE
        PS C:\> .\Get-Duplicates.ps1 -Items 1,2,3,2,4 -Unique
        1
        2
        3
        4

    .INPUTS
        System.Object[]

    .OUTPUTS
        System.Object[]

    .NOTES
    ===========================================================================
     Created with:     SAPIEN Technologies, Inc., PowerShell Studio 2014 v4.1.72
     Created on:       10/15/2014 9:34 AM
     Created by:       June Blender (juneb)
#>

param
(
    [Parameter(Mandatory = $true,
               ValueFromPipeline = $true)]
    [Object[]]
    $Items,
    
    [Parameter(Mandatory = $false)]
    [Switch]
    $Unique
)
Begin
{
    $hash = [ordered]@{ }
    $duplicates = @()
}
Process
{
    foreach ($item in $Items)
    {
        try
        {
            $hash.add($item, 0)
        }
        catch [System.Management.Automation.MethodInvocationException]
        {
            $duplicates += $item
        }
    }
}
End
{
    if ($unique)
    {
        return $hash.keys
        
    }
    elseif ($duplicates)
    {
        return $duplicates
    }
}
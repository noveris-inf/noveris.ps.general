<#
    .SYNOPSIS
        Get windows update status for systems
    .NOTES
        Author: Jesse Reichman (Noveris)
#>

[CmdletBinding(DefaultParameterSetName="retrieve")]
param(
    [Parameter(ParameterSetName="retrieve", Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,

    [Parameter(ParameterSetName="retrieve", Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$Filter,

    [Parameter(ParameterSetName="retrieve", Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$LDAPFilter,

    [Parameter(ParameterSetName="provided", Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Systems,

    [Parameter(ParameterSetName="retrieve", Mandatory=$false)]
    [ValidateNotNull()]
    [int]$MachineAge = 30,

    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [switch]$AsCSV = $false
)

########
# Global settings
Set-StrictMode -Version 2
$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"

<#
#>
Function Get-RemoteClassInstance
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClassName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName
    )

    process
    {
        #Attempt via CIM first
        try {
            Get-CimInstance -Property * -ComputerName $ComputerName -ClassName $ClassName
            return
        } catch {
            Write-Warning "Failed to retrieve information via CIM: $_"
        }

        # Fallback to WMI
        try {
            Get-WmiObject -Property * -ComputerName $ComputerName -Class $ClassName
            return
        } catch {
            Write-Warning "Failed to retrieve information via WMI: $_"
        }

        # Nothing worked, write-error
        Write-Error "No remaining methods to retrieve class information from system"
    }
}

########
# If required, retrieve a list of the systems using supplied parameters
if ($PSCmdlet.ParameterSetName -eq "retrieve")
{
    # Add relevant parameters that have been supplied
    $retrieve = @{
        Properties = "lastLogonDate"
    }
    "LDAPFilter", "Filter", "SearchBase" | ForEach-Object {
        if ($PSBoundParameters.Keys -contains $_)
        {
            $retrieve[$_] = $PSBoundParameters[$_]
        }
    }

    # Retrieve the list of systems
    try {
        $Systems = Get-ADComputer @retrieve |
            Where-Object { $_.lastLogonDate -gt [DateTime]::Now.AddDays(-[Math]::Abs($MachineAge)) } |
            ForEach-Object { $_.Name }
    } catch {
        Write-Information "Failed to retrieve a list of systems from active directory: $_"
        throw $_
    }
}

########
# Iterate through each system to get update details
$results = $Systems | ForEach-Object {
    $name = $_

    Write-Verbose "Operating on $name"
    $state = [PSCustomObject]@{
        System = $name
        Type = "Unknown"
        Version = "Unknown"
        Critical = -1
        Security = -1
        SecurityAge = -1
    }

    # Retrieve licensing information
    try {
        Write-Verbose "Retrieving update information"

        $block = {
            $updates = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher().Search("IsInstalled=0 and IsHidden=0").Updates
            $updates | ForEach-Object {
                [PSCustomObject]@{
                    Title = $_.Title
                    MsrcSeverity = $_.MsrcSeverity
                    Published = $_.LastDeploymentChangeTime
                    Security = [bool](($_.Categories | Where-Object { $_.Name -eq "Security Updates" } | Measure-Object).Count -gt 0)
                }
            }
        }

        $updates = Invoke-Command -ComputerName $name -ScriptBlock $block
        $state.Critical = ($updates | Where-Object { $_.MsrcSeverity -eq "Critical" } | Measure-Object).Count
        $state.Security = ($updates | Where-Object { $_.Security -eq $true } | Measure-Object).Count
        $oldest = ($updates | Where-Object { $_.Security -eq $true }) | Sort-Object -Property Published | Select-Object -First 1
        if ($oldest -ne $null)
        {
            $state.SecurityAge = [Math]::Round(([DateTime]::Now - $oldest.Published).TotalDays, 0)
        }
    } catch {
        Write-Warning "Failed to retrieve update information from ${name}: $_"
    }

    # Retrieve system information
    try {
        Write-Verbose "Retrieving system info"

        $sysinfo = Get-RemoteClassInstance -ComputerName $name -ClassName Win32_OperatingSystem
        $state.Type = $sysinfo.Caption
        $state.Version = $sysinfo.Version
    } catch {
        Write-Warning "Failed to retrieve system information from ${name}: $_"
    }

    $state
}

########
# Display output. Format as CSV, if requested
if ($AsCSV)
{
    $results | ConvertTo-CSV -NoTypeInformation
} else {
    $results
}

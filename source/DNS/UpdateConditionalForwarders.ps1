[CmdletBinding()]
param(
    [Parameter(mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Domain,

    [Parameter(mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ManualServers,

    [Parameter(mandatory=$false)]
    [switch]$IgnoreLocalResolver = $false,

    [Parameter(mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$LocalDNSServers,

    [Parameter(mandatory=$false)]
    [switch]$IgnoreExistingEntries = $false,

    [Parameter(mandatory=$false)]
    [ValidateNotNull()]
    [int]$EntryLimit = $null,

    [Parameter(mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogFile
)


################
# Global settings
Set-StrictMode -Version latest
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

# Clear log file
if (![string]::IsNullOrEmpty($LogFile))
{
    "" | Out-File -Encoding UTF8 $LogFile
}

& {
    # Seed list of servers for discovery
    $seedServers = New-Object 'System.Collections.Generic.HashSet[string]'

    ########
    # Add manual servers to the seed list
    Write-Information ""
    Write-Information "Adding any manually specified servers"
    $ManualServers | Where-Object { ![string]::IsNullOrEmpty($_) } | ForEach-Object {
        Write-Information "Manual seed server: $_"
        $seedServers.Add($_) | Out-Null
    }

    ########
    # Add name servers resolved using locally configured resolvers
    if (!$IgnoreLocalResolver)
    {
        Write-Information ""
        Write-Information "Using local resolver to determine authoritative name servers for ${Domain}"
        try {
            Resolve-DnsName -Type NS -Name $Domain |
                Where-Object {$_ -is 'Microsoft.DnsClient.Commands.DnsRecord_A' } |
                ForEach-Object {
                    Write-Information ("Automatic seed server: " + $_.IPAddress)
                    $seedServers.Add($_.IPAddress) | Out-Null
                }
        } catch {
            Write-Information "Failed to get NS entries using local resolver: $_"
        }
    }

    ########
    # Add servers determined from existing conditional forwarder
    if (!$IgnoreExistingEntries)
    {
        Write-Information ""
        Write-Information "Adding any existing conditional forwarders"
        foreach ($target in $LocalDnsServers)
        {
            try {
                $masterServers = (Get-DnsServerZone -Name $Domain).MasterServers
                $masterServers | ForEach-Object {
                    Write-Information "Existing entry ($target): $_"
                    $seedServers.Add($_.ToString()) | Out-Null
                }
            } catch {
                Write-Information "Failed to get conditional forwarder entries from ${target}: $_"
            }
        }
    }

    ########
    # Attempt to add any authoritative name servers to the seed list using first seed server
    # to return a NS list
    Write-Information ""
    Write-Information "Attempting to add any discovered authoritative name servers"
    foreach ($current in $seedServers)
    {
        try {
            $response = Resolve-DnsName -Name $Domain -Type NS -Server $current |
                Where-Object { $_ -is 'Microsoft.DnsClient.Commands.DnsRecord_A' } |
                ForEach-Object {
                    Write-Information ("Auth NS Server: " + $_.IPAddress)
                    $_.IPAddress
                }
        } catch {
            # Failed to resolve. Log error and continue to next server
            Write-Information "Failed to resolve using ${current}: $_"
            continue
        }

        if (($response | Measure-Object).Count -lt 1)
        {
            Write-Information "Empty NS server list from ${current}"
            continue
        }

        # Received something valid from this DNS server. Stop processing here
        #$server = $current
        #$authNameServers = $response
        $response | ForEach-Object { $seedServers.Add($_) | Out-Null }
        break
    }

    # Check that we now have a list of servers
    $resolverCount = $seedServers.Count
    if ($resolverCount -lt 1)
    {
        Write-Error "No seed servers identified automatically or specified manually. Servers can be specified manually with the ManualServers option."
    }

    $resolverCount = $seedServers.Count
    Write-Information ""
    Write-Information "Determined a list of ${resolverCount} servers for ${Domain}"
    Write-Information "Full server list:"
    $seedServers | Out-String

    # Determine reachability and responsivity for all DNS servers
    Write-Information ""
    Write-Information "Determining reachability and responsivity for name servers"
    $effectiveNameServers = $seedServers | ForEach-Object {
        $current = $_

        # Attempt to resolve nameservers for $Domain from server
        try {
            Write-Information "Testing server: $current"
            $start = [DateTime]::Now
            $response = Resolve-DnsName -Server $current -Type NS $Domain -QuickTimeout | Where-Object { $_ -is 'Microsoft.DnsClient.Commands.DnsRecord_A' } 
            $total = [DateTime]::Now - $start

            if (($response | Measure-Object).Count -gt 0)
            {
                [PSCustomObject]@{
                    Server = $current
                    ResponseTimeMS = $total.TotalMilliseconds
                }
            } else {
                Write-Information "No valid responses from $current"
            }
        } catch {
            Write-Information "Failed to resolve against ${current}: $_"
        }
    } | Sort-Object -Property ResponseTimeMS

    $resolverCount = ($effectiveNameServers | Measure-Object).Count
    Write-Information "Determined a list of ${resolverCount} reachable servers for ${Domain}"
    Write-Information "Full server list (ordered by response time):"
    $effectiveNameServers | Out-String

    # Limit the number of entries, if specified
    if ($EntryLimit -ne $null)
    {
        Write-Information "Limiting to top $EntryLimit fastest servers"
        $effectiveNameServers = $effectiveNameServers | Select-Object -First $EntryLimit
        $effectiveNameServers | Out-String
    }

    ########
    # Update the conditional forwarder entry
    if (($effectiveNameServers | Measure-Object).Count -lt 1)
    {
        Write-Information "Empty conditional forwarder list. Not updating servers."
    } else {
        foreach ($target in $LocalDNSServers)
        {
            Write-Information "Updating conditional forwarder for $Domain on $Target"
            try {
                Set-DnsServerConditionalForwarderZone -Name $Domain -MasterServers $effectiveNameServers.Server -ComputerName $target -Confirm:$false
            } catch {
                Write-Information "Failed to set conditional forwarder entries for ${target}: $_"
            }
        }
    }
} *>&1 | Out-String -Stream | ForEach-Object {
    Write-Information $_
    if (![string]::IsNullOrEmpty($LogFile))
    {
        $_ | Out-File -Encoding UTF8 $LogFile -Append
    }
}
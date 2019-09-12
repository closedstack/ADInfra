#Script to collect information for all DHCP Scopes on All DHCP Servers in the Domain
#Watch the debug.. scrolling text soothes anxiety whether the script is working
$DebugPreference = 'Continue'
Import-Module DhcpServer
$outputFile = "c:\reports\DHCPInventory.csv"
Function Get-DHCPServerScopes{
PARAM(
$DHCPServer
)

    $scopes = Get-DhcpServerv4Scope -ComputerName $DHCPServer
    $ServerOptions = Get-DhcpServerv4OptionValue -ComputerName $DHCPServer
    $SvrOption81 = $ServerOptions | where {$_.OptionId -eq 81}
    $SvrOption6 = $ServerOptions | where {$_.OptionId -eq 6}
    $results= @()
    foreach($scope in $scopes){
        Write-Debug "Options for $($scope.ScopeId)"
        $options = Get-DhcpServerv4OptionValue -ComputerName $DHCPServer -ScopeId $scope.ScopeId
        $DNSOptions = Get-DhcpServerv4DnsSetting -ComputerName $DHCPServer -ScopeId $($scope.ScopeId)
        #If there is no failover relationship for scope, it will throw and error, which is fine - we just have no value in that column
        $failover = Get-DhcpServerv4Failover -ScopeId $($scope.ScopeId) -ComputerName $DHCPServer -ErrorAction Continue
        $option51 = $options | where {$_.OptionId -eq 51}
        $option15 = $options | where {$_.OptionId -eq 15}
        $option6 = $options | where {$_.OptionId -eq 6}
        $hash = [ordered]@{
          DHCPServerName = "$DHCPServer"
          ServerOption81 = "$($SvrOption81.value)"
          ServerOption6 = "$($SvrOption6.value)"
          LeaseDuration = "$($option51.Value)"
          DnsDomainNameOption = "$($option15.value)"
          ScopeDNS = "$($option6.value)"
          DHCPDynamicUpdates = $DNSOptions.DynamicUpdates
          DeleteDnsRROnLeaseExpiry   = $DNSOptions.DeleteDnsRROnLeaseExpiry
          UpdateDnsRRForOlderClients = $DNSOptions.UpdateDnsRRForOlderClients
          DnsSuffix                  = $DNSOptions.DnsSuffix
          DisableDnsPtrRRUpdate      = $DNSOptions.DisableDnsPtrRRUpdate
          NameProtection             = $DNSOptions.NameProtection
          FailoverRelationShipName = "$($failover.Name)"
          FailoverPartner = "$($failover.PartnerServer)"
          FailoverMode = "$($failover.Mode)"
          FailoverServerRole = "$($failover.ServerRole)"
          FailoverEnableAuth = "$($failover.EnableAuth)"
          FailoverReservePercent = "$($failover.ReservePercent)"
          FailoverMCLT = "$($failover.MaxClientLeadTime)"
          FailoverStateSwitch = "$($failover.StateSwitchInterval)"
      }
        $results += New-Object -TypeName PSObject -Property $hash
    }
    return $results
}

$DHCPServers = Get-DhcpServerInDC 
$allDHCPScopes = @()
foreach($server in $DHCPServers){
    Write-Debug "Working on Server $($server.DnsName)"
    $allDHCPScopes += Get-DHCPServerScopes -DHCPServer $server.DnsName
}
Export-CSV -NoTypeInformation -Path $outputFile

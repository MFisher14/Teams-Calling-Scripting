# TeamsEmergencyLocations.psm1
# Module for collecting Teams emergency calling locations and policies

function Get-TeamsEmergencyLocationsData {
    <#
    .SYNOPSIS
        Collects all Teams emergency calling locations, policies, and configurations
    #>
    
    Write-Host "  → Collecting emergency locations..." -ForegroundColor Cyan
    
    $emergencyData = @{
        EmergencyCallingPolicies = @()
        EmergencyCallRoutingPolicies = @()
        EmergencyAddresses = @()
        EmergencyLocations = @()
        NetworkSites = @()
        TrustedIPAddresses = @()
        EmergencyNumbers = @()
        LocationInformationService = @()
        UserEmergencyAssignments = @()
    }
    
    try {
        # Emergency Calling Policies
        Write-Host "    - Emergency calling policies" -ForegroundColor DarkCyan
        $emergencyData.EmergencyCallingPolicies = @(Get-CsTeamsEmergencyCallingPolicy -ErrorAction SilentlyContinue)
        
        # Emergency Call Routing Policies
        Write-Host "    - Emergency call routing policies" -ForegroundColor DarkCyan
        $emergencyData.EmergencyCallRoutingPolicies = @(Get-CsTeamsEmergencyCallRoutingPolicy -ErrorAction SilentlyContinue)
        
        # Emergency Addresses
        Write-Host "    - Emergency addresses" -ForegroundColor DarkCyan
        try {
            $emergencyAddresses = @(Get-CsOnlineLisLocation -ErrorAction SilentlyContinue)
            $emergencyData.EmergencyAddresses = $emergencyAddresses
            
            # If we have addresses, try to get detailed location information
            if ($emergencyAddresses.Count -gt 0) {
                Write-Host "    - Processing $($emergencyAddresses.Count) emergency addresses" -ForegroundColor DarkCyan
                $detailedLocations = @()
                
                foreach ($address in $emergencyAddresses) {
                    try {
                        $locationDetail = @{
                            LocationId = $address.LocationId
                            Address = $address.Address
                            City = $address.City
                            CountryOrRegion = $address.CountryOrRegion
                            StateOrProvince = $address.StateOrProvince
                            PostalCode = $address.PostalCode
                            Description = $address.Description
                            CompanyName = $address.CompanyName
                            CompanyTaxId = $address.CompanyTaxId
                            HouseNumber = $address.HouseNumber
                            HouseNumberSuffix = $address.HouseNumberSuffix
                            PreDirectional = $address.PreDirectional
                            StreetName = $address.StreetName
                            StreetSuffix = $address.StreetSuffix
                            PostDirectional = $address.PostDirectional
                            Latitude = $address.Latitude
                            Longitude = $address.Longitude
                            Elin = $address.Elin
                            IsValidated = $address.IsValidated
                            ValidationStatus = $address.ValidationStatus
                        }
                        $detailedLocations += $locationDetail
                    }
                    catch {
                        Write-Verbose "Could not get detailed info for address: $($address.LocationId)"
                    }
                }
                $emergencyData.EmergencyLocations = $detailedLocations
            }
        }
        catch {
            Write-Warning "Could not retrieve emergency addresses: $($_.Exception.Message)"
            $emergencyData.EmergencyAddresses = @()
            $emergencyData.EmergencyLocations = @()
        }
        
        # Network Sites (related to emergency locations)
        Write-Host "    - Network sites" -ForegroundColor DarkCyan
        try {
            $networkSites = @(Get-CsTenantNetworkSite -ErrorAction SilentlyContinue)
            $emergencyData.NetworkSites = $networkSites
        }
        catch {
            Write-Verbose "Could not retrieve network sites"
            $emergencyData.NetworkSites = @()
        }
        
        # Trusted IP Addresses
        Write-Host "    - Trusted IP addresses" -ForegroundColor DarkCyan
        try {
            $trustedIPs = @(Get-CsTenantTrustedIPAddress -ErrorAction SilentlyContinue)
            $emergencyData.TrustedIPAddresses = $trustedIPs
        }
        catch {
            Write-Verbose "Could not retrieve trusted IP addresses"
            $emergencyData.TrustedIPAddresses = @()
        }
        
        # Emergency Numbers Configuration
        Write-Host "    - Emergency numbers" -ForegroundColor DarkCyan
        try {
            $emergencyNumbers = @()
            
            # Try to get emergency numbers from policies
            foreach ($policy in $emergencyData.EmergencyCallRoutingPolicies) {
                if ($policy.EmergencyNumbers) {
                    foreach ($number in $policy.EmergencyNumbers) {
                        $emergencyNumbers += @{
                            PolicyName = $policy.Identity
                            EmergencyDialString = $number.EmergencyDialString
                            EmergencyDialMask = $number.EmergencyDialMask
                            OnlinePSTNUsage = $number.OnlinePSTNUsage
                        }
                    }
                }
            }
            $emergencyData.EmergencyNumbers = $emergencyNumbers
        }
        catch {
            Write-Verbose "Could not retrieve emergency numbers configuration"
            $emergencyData.EmergencyNumbers = @()
        }
        
        # Location Information Service Configuration
        Write-Host "    - Location Information Service" -ForegroundColor DarkCyan
        try {
            $lisConfig = @()
            
            # Try to get LIS configuration information
            $lisWirelessAccessPoints = @(Get-CsOnlineLisWirelessAccessPoint -ErrorAction SilentlyContinue)
            $lisSubnets = @(Get-CsOnlineLisSubnet -ErrorAction SilentlyContinue)
            $lisPorts = @(Get-CsOnlineLisPort -ErrorAction SilentlyContinue)
            $lisSwitches = @(Get-CsOnlineLisSwitch -ErrorAction SilentlyContinue)
            
            $lisConfig += @{
                WirelessAccessPoints = $lisWirelessAccessPoints
                Subnets = $lisSubnets
                Ports = $lisPorts
                Switches = $lisSwitches
            }
            
            $emergencyData.LocationInformationService = $lisConfig
        }
        catch {
            Write-Verbose "Could not retrieve Location Information Service configuration"
            $emergencyData.LocationInformationService = @()
        }
        
        # User Emergency Policy Assignments (summary)
        Write-Host "    - User emergency policy assignments" -ForegroundColor DarkCyan
        try {
            $userAssignments = @()
            
            # Get users with emergency policies assigned
            $usersWithEmergencyPolicies = @(Get-CsOnlineUser -ErrorAction SilentlyContinue | Where-Object { 
                $null -ne $_.EmergencyCallingPolicy -or $null -ne $_.EmergencyCallRoutingPolicy -or $null -ne $_.LocationId 
            })
            
            foreach ($user in $usersWithEmergencyPolicies) {
                $assignment = @{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    EmergencyCallingPolicy = $user.EmergencyCallingPolicy
                    EmergencyCallRoutingPolicy = $user.EmergencyCallRoutingPolicy
                    LocationId = $user.LocationId
                    UsageLocation = $user.UsageLocation
                    Country = $user.Country
                    City = $user.City
                    StateOrProvince = $user.StateOrProvince
                    PostalCode = $user.PostalCode
                }
                $userAssignments += $assignment
            }
            
            $emergencyData.UserEmergencyAssignments = $userAssignments
        }
        catch {
            Write-Warning "Could not retrieve user emergency policy assignments: $($_.Exception.Message)"
            $emergencyData.UserEmergencyAssignments = @()
        }
        
        Write-Host "    ✓ Emergency locations collection complete" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error collecting emergency locations data: $($_.Exception.Message)"
    }
    
    return $emergencyData
}

Export-ModuleMember -Function Get-TeamsEmergencyLocationsData
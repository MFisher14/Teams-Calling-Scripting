# TeamsCallingPolicies.psm1
# Module f        # "CallingRegions" = try { Get-CsTenantCallingRegion -ErrorAction Stop } catch { Write-Warning "Could not retrieve Tenant Calling Regions. The cmdlet 'Get-CsTenantCallingRegion' may be obsolete or require different permissions."; $null }r collecting Teams calling policies and configurations

function Get-TeamsCallingPoliciesData {
    <#
    .SYNOPSIS
        Collects all Teams calling policies and related configurations
    #>
    
    Write-Host "  → Collecting calling policies..." -ForegroundColor Cyan
    
    $policiesData = @{
        TeamsCallingPolicies = @()
        VoiceApplications = @()
        TeamsIPPhonePolicies = @()
        TeamsCallingRegion = @()
        TeamsBranchSurvivabilityPolicy = @()
        TeamsDialoutPolicy = @()
        TeamsNetworkConfiguration = @()
        CallingLineIdentityPolicies = @()
        PrivacyPolicies = @()
    }
    
    try {
        # Teams Calling Policies
        Write-Host "    - Teams calling policies" -ForegroundColor DarkCyan
        $policiesData.TeamsCallingPolicies = @(Get-CsTeamsCallingPolicy -ErrorAction SilentlyContinue)
        
        # Voice Applications
        Write-Host "    - Voice applications" -ForegroundColor DarkCyan
        $policiesData.VoiceApplications = @(Get-CsApplicationAccessPolicy -ErrorAction SilentlyContinue)
        
        # IP Phone Policies
        Write-Host "    - IP phone policies" -ForegroundColor DarkCyan
        $policiesData.TeamsIPPhonePolicies = @(Get-CsTeamsIPPhonePolicy -ErrorAction SilentlyContinue)
        
        # Calling Regions
        Write-Host "    - Calling regions" -ForegroundColor DarkCyan
                # "CallingRegions" = try { Get-CsTenantCallingRegion -ErrorAction Stop } catch { Write-Warning "Could not retrieve Tenant Calling Regions. The cmdlet 'Get-CsTenantCallingRegion' may be obsolete or require different permissions."; $null }
        
        # Branch Survivability Policies
        Write-Host "    - Branch survivability policies" -ForegroundColor DarkCyan
        # Note: Get-CsTeamsBranchSurvivabilityPolicy may be obsolete in newer versions
        try {
            $policiesData.TeamsBranchSurvivabilityPolicy = @(Get-CsTeamsBranchSurvivabilityPolicy -ErrorAction Stop)
        }
        catch {
            Write-Warning "Could not retrieve Branch Survivability Policies. The cmdlet 'Get-CsTeamsBranchSurvivabilityPolicy' may be obsolete or require different permissions."
            $policiesData.TeamsBranchSurvivabilityPolicy = @()
        }
        
        # Dialout Policies
        Write-Host "    - Dialout policies" -ForegroundColor DarkCyan
        # Note: Get-CsTeamsDialoutPolicy may be obsolete in newer versions
        try {
            $policiesData.TeamsDialoutPolicy = @(Get-CsTeamsDialoutPolicy -ErrorAction Stop)
        }
        catch {
            Write-Warning "Could not retrieve Teams Dialout Policies. The cmdlet 'Get-CsTeamsDialoutPolicy' may be obsolete or require different permissions."
            $policiesData.TeamsDialoutPolicy = @()
        }
        
        # Network Configuration
        Write-Host "    - Network configuration" -ForegroundColor DarkCyan
        try {
            $networkSites = @(Get-CsTenantNetworkSite -ErrorAction SilentlyContinue)
            $networkRegions = @(Get-CsTenantNetworkRegion -ErrorAction SilentlyContinue)
            $networkSubnets = @(Get-CsTenantNetworkSubnet -ErrorAction SilentlyContinue)
            
            $policiesData.TeamsNetworkConfiguration = @{
                NetworkSites = $networkSites
                NetworkRegions = $networkRegions
                NetworkSubnets = $networkSubnets
            }
        }
        catch {
            Write-Warning "Failed to collect network configuration: $($_.Exception.Message)"
        }
        
        # Calling Line Identity Policies
        Write-Host "    - Calling line identity policies" -ForegroundColor DarkCyan
        $policiesData.CallingLineIdentityPolicies = @(Get-CsCallingLineIdentity -ErrorAction SilentlyContinue)
        
        # Privacy Policies
        Write-Host "    - Privacy policies" -ForegroundColor DarkCyan
        $policiesData.PrivacyPolicies = @(Get-CsPrivacyConfiguration -ErrorAction SilentlyContinue)
        
        Write-Host "    ✓ Calling policies collection complete" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error collecting calling policies: $($_.Exception.Message)"
    }
    
    return $policiesData
}

Export-ModuleMember -Function Get-TeamsCallingPoliciesData
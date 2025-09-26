# TeamsCallingPolicies.psm1
# Module for collecting Teams calling policies and configurations

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
        $policiesData.VoiceApplications = @(Get-CsVoiceApplicationsPolicy -ErrorAction SilentlyContinue)
        
        # IP Phone Policies
        Write-Host "    - IP phone policies" -ForegroundColor DarkCyan
        $policiesData.TeamsIPPhonePolicies = @(Get-CsTeamsIPPhonePolicy -ErrorAction SilentlyContinue)
        
        # Calling Regions
        Write-Host "    - Calling regions" -ForegroundColor DarkCyan
        $policiesData.TeamsCallingRegion = @(Get-CsTenantCallingRegion -ErrorAction SilentlyContinue)
        
        # Branch Survivability Policies
        Write-Host "    - Branch survivability policies" -ForegroundColor DarkCyan
        $policiesData.TeamsBranchSurvivabilityPolicy = @(Get-CsTeamsBranchSurvivabilityPolicy -ErrorAction SilentlyContinue)
        
        # Dialout Policies
        Write-Host "    - Dialout policies" -ForegroundColor DarkCyan
        $policiesData.TeamsDialoutPolicy = @(Get-CsTeamsDialoutPolicy -ErrorAction SilentlyContinue)
        
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
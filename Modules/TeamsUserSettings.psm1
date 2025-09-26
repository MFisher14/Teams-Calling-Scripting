# TeamsUserSettings.psm1
# Module for collecting Teams user calling settings and assignments

function Get-TeamsUserSettingsData {
    <#
    .SYNOPSIS
        Collects all Teams user calling settings and policy assignments
    .PARAMETER IncludeDetailedData
        Switch to include detailed per-user data (may take longer for large tenants)
    #>
    param(
        [switch]$IncludeDetailedData
    )
    
    Write-Host "  → Collecting user settings..." -ForegroundColor Cyan
    
    $userSettingsData = @{
        UserPolicyAssignments = @()
        VoiceUserSettings = @()
        UserCallingSettings = @()
        LineURIAssignments = @()
        UserLocationAssignments = @()
        PolicyAssignmentsSummary = @()
        UsersWithCallingPlan = @()
        UsersWithDirectRouting = @()
        OnlineVoiceUsers = @()
    }
    
    try {
        # Get basic user information with voice capabilities
        Write-Host "    - Online voice users" -ForegroundColor DarkCyan
        $onlineVoiceUsers = @(Get-CsOnlineUser -ErrorAction SilentlyContinue | Where-Object { 
            $_.EnterpriseVoiceEnabled -eq $true -or 
            $_.LineURI -ne $null -or 
            $_.OnPremLineURI -ne $null -or
            $_.VoicePolicy -ne $null
        })
        $userSettingsData.OnlineVoiceUsers = $onlineVoiceUsers
        
        if ($onlineVoiceUsers.Count -gt 0) {
            Write-Host "    - Processing $($onlineVoiceUsers.Count) voice-enabled users" -ForegroundColor DarkCyan
            
            # Collect summary statistics
            $policyAssignmentsSummary = @()
            $usersWithCallingPlan = @()
            $usersWithDirectRouting = @()
            $userPolicyAssignments = @()
            $voiceUserSettings = @()
            $lineUriAssignments = @()
            $locationAssignments = @()
            
            foreach ($user in $onlineVoiceUsers) {
                try {
                    # Basic user voice settings
                    $userVoiceInfo = @{
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        EnterpriseVoiceEnabled = $user.EnterpriseVoiceEnabled
                        LineURI = $user.LineURI
                        OnPremLineURI = $user.OnPremLineURI
                        VoicePolicy = $user.VoicePolicy
                        VoiceRoutingPolicy = $user.VoiceRoutingPolicy
                        OnlineVoiceRoutingPolicy = $user.OnlineVoiceRoutingPolicy
                        CallingLineIdentity = $user.CallingLineIdentity
                        UsageLocation = $user.UsageLocation
                        Country = $user.Country
                        City = $user.City
                        Department = $user.Department
                        Title = $user.Title
                        AccountType = $user.AccountType
                        InterpretedUserType = $user.InterpretedUserType
                    }
                    
                    $voiceUserSettings += $userVoiceInfo
                    
                    # Determine calling plan vs direct routing
                    if ($user.LineURI -match "^tel:\+") {
                        if ($user.VoicePolicy -eq "BusinessVoice" -or $user.VoicePolicy -eq "HybridVoice") {
                            $usersWithCallingPlan += $userVoiceInfo
                        }
                        elseif ($user.OnlineVoiceRoutingPolicy -ne $null -or $user.VoiceRoutingPolicy -ne $null) {
                            $usersWithDirectRouting += $userVoiceInfo
                        }
                    }
                    
                    # Line URI assignments
                    if ($user.LineURI -ne $null -or $user.OnPremLineURI -ne $null) {
                        $lineUriInfo = @{
                            UserPrincipalName = $user.UserPrincipalName
                            DisplayName = $user.DisplayName
                            LineURI = $user.LineURI
                            OnPremLineURI = $user.OnPremLineURI
                            PrivateLine = $user.PrivateLine
                            LineURISource = if ($user.LineURI -ne $null) { "Online" } elseif ($user.OnPremLineURI -ne $null) { "OnPrem" } else { "None" }
                        }
                        $lineUriAssignments += $lineUriInfo
                    }
                    
                    # If detailed data is requested, get policy assignments
                    if ($IncludeDetailedData) {
                        try {
                            # Get detailed policy assignments for this user
                            $policyAssignment = @{
                                UserPrincipalName = $user.UserPrincipalName
                                DisplayName = $user.DisplayName
                                TeamsCallingPolicy = $user.TeamsCallingPolicy
                                TeamsCallHoldPolicy = $user.TeamsCallHoldPolicy
                                TeamsCallParkPolicy = $user.TeamsCallParkPolicy
                                TeamsCortanaPolicy = $user.TeamsCortanaPolicy
                                TeamsIPPhonePolicy = $user.TeamsIPPhonePolicy
                                TeamsMobilityPolicy = $user.TeamsMobilityPolicy
                                TeamsVoiceApplicationsPolicy = $user.TeamsVoiceApplicationsPolicy
                                CallingLineIdentity = $user.CallingLineIdentity
                                OnlineDialinConferencingPolicy = $user.OnlineDialinConferencingPolicy
                                OnlineVoicemailPolicy = $user.OnlineVoicemailPolicy
                                OnlineVoiceRoutingPolicy = $user.OnlineVoiceRoutingPolicy
                                TenantDialPlan = $user.TenantDialPlan
                            }
                            $userPolicyAssignments += $policyAssignment
                            
                            # Try to get user location information
                            try {
                                $locationInfo = @{
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
                                    CompanyName = $user.CompanyName
                                    OnPremisesExtensionAttributes = if ($user.OnPremisesExtensionAttributes) {
                                        @{
                                            extensionAttribute1 = $user.OnPremisesExtensionAttributes.extensionAttribute1
                                            extensionAttribute2 = $user.OnPremisesExtensionAttributes.extensionAttribute2
                                            extensionAttribute3 = $user.OnPremisesExtensionAttributes.extensionAttribute3
                                        }
                                    } else { $null }
                                }
                                $locationAssignments += $locationInfo
                            }
                            catch {
                                Write-Verbose "Could not get location info for user: $($user.UserPrincipalName)"
                            }
                        }
                        catch {
                            Write-Verbose "Could not get detailed policy assignments for user: $($user.UserPrincipalName)"
                        }
                    }
                }
                catch {
                    Write-Verbose "Error processing user $($user.UserPrincipalName): $($_.Exception.Message)"
                }
            }
            
            # Create policy assignments summary
            if ($userPolicyAssignments.Count -gt 0) {
                # Group by policy types to create summary
                $policyTypes = @(
                    "TeamsCallingPolicy", "TeamsCallHoldPolicy", "TeamsCallParkPolicy", 
                    "TeamsCortanaPolicy", "TeamsIPPhonePolicy", "TeamsMobilityPolicy",
                    "TeamsVoiceApplicationsPolicy", "CallingLineIdentity", "OnlineDialinConferencingPolicy",
                    "OnlineVoicemailPolicy", "OnlineVoiceRoutingPolicy", "TenantDialPlan"
                )
                
                foreach ($policyType in $policyTypes) {
                    $policyGroups = $userPolicyAssignments | Group-Object $policyType | ForEach-Object {
                        @{
                            PolicyValue = $_.Name
                            UserCount = $_.Count
                            Users = @($_.Group | ForEach-Object { $_.UserPrincipalName })
                        }
                    }
                    
                    $policyAssignmentsSummary += @{
                        PolicyType = $policyType
                        TotalUsers = $userPolicyAssignments.Count
                        PolicyDistribution = $policyGroups
                    }
                }
            }
            
            $userSettingsData.UserPolicyAssignments = $userPolicyAssignments
            $userSettingsData.VoiceUserSettings = $voiceUserSettings
            $userSettingsData.LineURIAssignments = $lineUriAssignments
            $userSettingsData.UserLocationAssignments = $locationAssignments
            $userSettingsData.PolicyAssignmentsSummary = $policyAssignmentsSummary
            $userSettingsData.UsersWithCallingPlan = $usersWithCallingPlan
            $userSettingsData.UsersWithDirectRouting = $usersWithDirectRouting
        }
        
        # Get calling settings for voice users (if not already collected in detail)
        if (-not $IncludeDetailedData -and $onlineVoiceUsers.Count -gt 0) {
            Write-Host "    - User calling settings summary" -ForegroundColor DarkCyan
            $callingSettings = @()
            
            # Create a summary of calling settings without detailed per-user data
            $voiceEnabledCount = ($onlineVoiceUsers | Where-Object { $_.EnterpriseVoiceEnabled -eq $true }).Count
            $lineUriCount = ($onlineVoiceUsers | Where-Object { $_.LineURI -ne $null -or $_.OnPremLineURI -ne $null }).Count
            $callingPlanCount = ($onlineVoiceUsers | Where-Object { $_.VoicePolicy -eq "BusinessVoice" -or $_.VoicePolicy -eq "HybridVoice" }).Count
            $directRoutingCount = ($onlineVoiceUsers | Where-Object { $_.OnlineVoiceRoutingPolicy -ne $null -or $_.VoiceRoutingPolicy -ne $null }).Count
            
            $callingSettings += @{
                TotalUsers = $onlineVoiceUsers.Count
                EnterpriseVoiceEnabled = $voiceEnabledCount
                UsersWithLineURI = $lineUriCount
                UsersWithCallingPlan = $callingPlanCount
                UsersWithDirectRouting = $directRoutingCount
                CollectedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
            }
            
            $userSettingsData.UserCallingSettings = $callingSettings
        }
        
        Write-Host "    ✓ User settings collection complete ($($onlineVoiceUsers.Count) voice users)" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error collecting user settings: $($_.Exception.Message)"
    }
    
    return $userSettingsData
}

Export-ModuleMember -Function Get-TeamsUserSettingsData
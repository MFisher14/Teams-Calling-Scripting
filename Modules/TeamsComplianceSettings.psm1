# TeamsComplianceSettings.psm1
# Module for collecting Teams compliance, security, and audit settings

function Get-TeamsComplianceSettingsData {
    <#
    .SYNOPSIS
        Collects all Teams compliance, security, and audit configurations
    #>
    
    Write-Host "  → Collecting compliance settings..." -ForegroundColor Cyan
    
    $complianceData = @{
        ComplianceRecordingPolicies = @()
        RetentionPolicies = @()
        CallRecordingPolicies = @()
        InformationBarrierPolicies = @()
        DLPPolicies = @()
        AuditSettings = @()
        SecuritySettings = @()
        PrivacySettings = @()
        RecordingApplications = @()
        ComplianceNotifications = @()
        DataRetentionSettings = @()
        CallAnalyticsSettings = @()
    }
    
    try {
        # Compliance Recording Policies
        Write-Host "    - Compliance recording policies" -ForegroundColor DarkCyan
        try {
            $complianceRecordingPolicies = @(Get-CsTeamsComplianceRecordingPolicy -ErrorAction SilentlyContinue)
            $complianceData.ComplianceRecordingPolicies = $complianceRecordingPolicies
        }
        catch {
            Write-Verbose "Compliance recording policies not available"
            $complianceData.ComplianceRecordingPolicies = @()
        }
        
        # Call Recording Policies
        Write-Host "    - Call recording policies" -ForegroundColor DarkCyan
        try {
            $callRecordingPolicies = @(Get-CsTeamsCallRecordingPolicy -ErrorAction SilentlyContinue)
            $complianceData.CallRecordingPolicies = $callRecordingPolicies
        }
        catch {
            Write-Verbose "Call recording policies not available"
            $complianceData.CallRecordingPolicies = @()
        }
        
        # Recording Applications
        Write-Host "    - Recording applications" -ForegroundColor DarkCyan
        try {
            $recordingApplications = @(Get-CsTeamsComplianceRecordingApplication -ErrorAction SilentlyContinue)
            $complianceData.RecordingApplications = $recordingApplications
        }
        catch {
            Write-Verbose "Recording applications not available"
            $complianceData.RecordingApplications = @()
        }
        
        # Information Barrier Policies
        Write-Host "    - Information barrier policies" -ForegroundColor DarkCyan
        try {
            # Note: This requires Security & Compliance PowerShell module
            $ibPolicies = @(Get-InformationBarrierPolicy -ErrorAction SilentlyContinue)
            $complianceData.InformationBarrierPolicies = $ibPolicies
        }
        catch {
            Write-Verbose "Information barrier policies not available (requires Security & Compliance module)"
            $complianceData.InformationBarrierPolicies = @()
        }
        
        # Security Settings
        Write-Host "    - Security settings" -ForegroundColor DarkCyan
        try {
            $securitySettings = @{
                TenantFederationSettings = $null
                ExternalAccessPolicy = $null
                ClientPolicy = $null
                ConferencingPolicy = $null
                PrivacyConfiguration = $null
            }
            
            # Try to get various security-related configurations
            try {
                $securitySettings.TenantFederationSettings = Get-CsTenantFederationConfiguration -ErrorAction SilentlyContinue
                $securitySettings.ExternalAccessPolicy = @(Get-CsExternalAccessPolicy -ErrorAction SilentlyContinue)
                $securitySettings.ClientPolicy = @(Get-CsClientPolicy -ErrorAction SilentlyContinue)
                $securitySettings.ConferencingPolicy = @(Get-CsConferencingPolicy -ErrorAction SilentlyContinue)
                $securitySettings.PrivacyConfiguration = Get-CsPrivacyConfiguration -ErrorAction SilentlyContinue
            }
            catch {
                Write-Verbose "Some security settings not available"
            }
            
            $complianceData.SecuritySettings = $securitySettings
        }
        catch {
            Write-Verbose "Security settings not available"
            $complianceData.SecuritySettings = @()
        }
        
        # Privacy Settings
        Write-Host "    - Privacy settings" -ForegroundColor DarkCyan
        try {
            $privacySettings = @{
                TeamsPrivacyConfiguration = $null
                CallerIdPolicies = @()
                PrivacyPolicies = @()
            }
            
            # Teams Privacy Configuration
            $privacySettings.TeamsPrivacyConfiguration = Get-CsTeamsPrivacyConfiguration -ErrorAction SilentlyContinue
            
            # Caller ID Policies
            $privacySettings.CallerIdPolicies = @(Get-CsCallerIdPolicy -ErrorAction SilentlyContinue)
            
            # Privacy Policies
            $privacySettings.PrivacyPolicies = @(Get-CsPrivacyConfiguration -ErrorAction SilentlyContinue)
            
            $complianceData.PrivacySettings = $privacySettings
        }
        catch {
            Write-Verbose "Privacy settings not available"
            $complianceData.PrivacySettings = @()
        }
        
        # Audit Settings
        Write-Host "    - Audit settings" -ForegroundColor DarkCyan
        try {
            $auditSettings = @{
                AuditConfiguration = $null
                TeamsAuditConfiguration = $null
                AdminAuditLogConfig = $null
            }
            
            # Try to get audit configurations
            try {
                $auditSettings.AuditConfiguration = Get-AdminAuditLogConfig -ErrorAction SilentlyContinue
                $auditSettings.TeamsAuditConfiguration = Get-CsTeamsCallingPolicy -Identity Global -ErrorAction SilentlyContinue | Select-Object -Property *Audit*
            }
            catch {
                Write-Verbose "Some audit settings not available"
            }
            
            $complianceData.AuditSettings = $auditSettings
        }
        catch {
            Write-Verbose "Audit settings not available"
            $complianceData.AuditSettings = @()
        }
        
        # Data Retention Settings
        Write-Host "    - Data retention settings" -ForegroundColor DarkCyan
        try {
            $retentionSettings = @{
                TeamsRetentionPolicy = @()
                MessagingPolicy = @()
                MeetingPolicy = @()
            }
            
            # Teams Retention Policies (if available)
            try {
                $retentionSettings.TeamsRetentionPolicy = @(Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue | Where-Object { 
                    $_.Workload -contains "Teams" -or $_.ExchangeLocation -contains "All" 
                })
            }
            catch {
                Write-Verbose "Retention policies not available (requires Security & Compliance module)"
            }
            
            # Teams Messaging Policies (retention-related settings)
            $retentionSettings.MessagingPolicy = @(Get-CsTeamsMessagingPolicy -ErrorAction SilentlyContinue)
            
            # Teams Meeting Policies (retention-related settings)
            $retentionSettings.MeetingPolicy = @(Get-CsTeamsMeetingPolicy -ErrorAction SilentlyContinue)
            
            $complianceData.DataRetentionSettings = $retentionSettings
        }
        catch {
            Write-Verbose "Data retention settings not available"
            $complianceData.DataRetentionSettings = @()
        }
        
        # Call Analytics Settings
        Write-Host "    - Call analytics settings" -ForegroundColor DarkCyan
        try {
            $analyticsSettings = @{
                CallAnalyticsConfiguration = $null
                CallQualityDashboard = $null
                TeamsAnalyticsPolicy = @()
            }
            
            # Try to get call analytics configurations
            try {
                # Note: These may require specific permissions or modules
                $analyticsSettings.CallAnalyticsConfiguration = Get-CsCallAnalyticsConfiguration -ErrorAction SilentlyContinue
                $analyticsSettings.TeamsAnalyticsPolicy = @(Get-CsTeamsAnalyticsPolicy -ErrorAction SilentlyContinue)
            }
            catch {
                Write-Verbose "Call analytics settings not available"
            }
            
            $complianceData.CallAnalyticsSettings = $analyticsSettings
        }
        catch {
            Write-Verbose "Call analytics settings not available"
            $complianceData.CallAnalyticsSettings = @()
        }
        
        # Compliance Notifications
        Write-Host "    - Compliance notifications" -ForegroundColor DarkCyan
        try {
            $notificationSettings = @{
                ComplianceNotificationPolicies = @()
                AlertPolicies = @()
            }
            
            # Try to get notification policies (if available)
            try {
                # This may require Security & Compliance PowerShell
                $notificationSettings.AlertPolicies = @(Get-ProtectionAlert -ErrorAction SilentlyContinue)
            }
            catch {
                Write-Verbose "Compliance notifications not available"
            }
            
            $complianceData.ComplianceNotifications = $notificationSettings
        }
        catch {
            Write-Verbose "Compliance notifications not available"
            $complianceData.ComplianceNotifications = @()
        }
        
        Write-Host "    ✓ Compliance settings collection complete" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error collecting compliance settings: $($_.Exception.Message)"
    }
    
    return $complianceData
}

Export-ModuleMember -Function Get-TeamsComplianceSettingsData
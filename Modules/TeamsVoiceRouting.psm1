# TeamsVoiceRouting.psm1
# Module for collecting Teams voice routing and PSTN configurations

function Get-TeamsVoiceRoutingData {
    <#
    .SYNOPSIS
        Collects all Teams voice routing, PSTN, and telephony configurations
    #>
    
    Write-Host "  → Collecting voice routing..." -ForegroundColor Cyan
    
    $voiceRoutingData = @{
        VoiceRoutingPolicies = @()
        VoiceRoutes = @()
        PSTNUsages = @()
        TenantDialPlans = @()
        SBCConfigurations = @()
        PSTNGateways = @()
        VoiceNormalizationRules = @()
        TrunkConfigurations = @()
        MediaConfigurations = @()
        CallRouting = @()
        NumberPorting = @()
        PhoneNumberAssignments = @()
    }
    
    try {
        # Voice Routing Policies
        Write-Host "    - Voice routing policies" -ForegroundColor DarkCyan
        $voiceRoutingData.VoiceRoutingPolicies = @(Get-CsOnlineVoiceRoutingPolicy -ErrorAction SilentlyContinue)
        
        # Voice Routes
        Write-Host "    - Voice routes" -ForegroundColor DarkCyan
        $voiceRoutingData.VoiceRoutes = @(Get-CsOnlineVoiceRoute -ErrorAction SilentlyContinue)
        
        # PSTN Usages
        Write-Host "    - PSTN usages" -ForegroundColor DarkCyan
        $voiceRoutingData.PSTNUsages = @(Get-CsOnlinePstnUsage -ErrorAction SilentlyContinue)
        
        # Tenant Dial Plans
        Write-Host "    - Tenant dial plans" -ForegroundColor DarkCyan
        $tenantDialPlans = @(Get-CsTenantDialPlan -ErrorAction SilentlyContinue)
        $voiceRoutingData.TenantDialPlans = $tenantDialPlans
        
        # Voice Normalization Rules (from dial plans)
        if ($tenantDialPlans.Count -gt 0) {
            Write-Host "    - Voice normalization rules" -ForegroundColor DarkCyan
            $normalizationRules = @()
            
            foreach ($dialPlan in $tenantDialPlans) {
                if ($dialPlan.NormalizationRules) {
                    foreach ($rule in $dialPlan.NormalizationRules) {
                        $normalizationRules += @{
                            DialPlanName = $dialPlan.Identity
                            RuleName = $rule.Name
                            Description = $rule.Description
                            Pattern = $rule.Pattern
                            Translation = $rule.Translation
                            Priority = $rule.Priority
                            IsInternalExtension = $rule.IsInternalExtension
                        }
                    }
                }
            }
            $voiceRoutingData.VoiceNormalizationRules = $normalizationRules
        }
        
        # Session Border Controllers (SBCs) - Direct Routing
        Write-Host "    - Session Border Controllers" -ForegroundColor DarkCyan
        try {
            $sbcConfigurations = @(Get-CsOnlinePSTNGateway -ErrorAction SilentlyContinue)
            $voiceRoutingData.SBCConfigurations = $sbcConfigurations
        }
        catch {
            Write-Verbose "Could not retrieve SBC configurations"
            $voiceRoutingData.SBCConfigurations = @()
        }
        
        # PSTN Gateways (legacy)
        Write-Host "    - PSTN gateways" -ForegroundColor DarkCyan
        try {
            $pstnGateways = @(Get-CsPstnGateway -ErrorAction SilentlyContinue)
            $voiceRoutingData.PSTNGateways = $pstnGateways
        }
        catch {
            Write-Verbose "PSTN gateways not available"
            $voiceRoutingData.PSTNGateways = @()
        }
        
        # Trunk Configurations
        Write-Host "    - Trunk configurations" -ForegroundColor DarkCyan
        try {
            $trunkConfigurations = @(Get-CsTrunkConfiguration -ErrorAction SilentlyContinue)
            $voiceRoutingData.TrunkConfigurations = $trunkConfigurations
        }
        catch {
            Write-Verbose "Trunk configurations not available"
            $voiceRoutingData.TrunkConfigurations = @()
        }
        
        # Media Configurations
        Write-Host "    - Media configurations" -ForegroundColor DarkCyan
        try {
            $mediaConfigurations = @(Get-CsMediaConfiguration -ErrorAction SilentlyContinue)
            $voiceRoutingData.MediaConfigurations = $mediaConfigurations
        }
        catch {
            Write-Verbose "Media configurations not available"
            $voiceRoutingData.MediaConfigurations = @()
        }
        
        # Call Routing Summary
        Write-Host "    - Call routing analysis" -ForegroundColor DarkCyan
        try {
            $callRoutingAnalysis = @()
            
            # Analyze voice routes and their PSTN usage relationships
            foreach ($route in $voiceRoutingData.VoiceRoutes) {
                $routeAnalysis = @{
                    RouteName = $route.Identity
                    NumberPattern = $route.NumberPattern
                    PstnUsages = if ($route.OnlinePstnUsages) { $route.OnlinePstnUsages } else { @() }
                    PstnGateways = if ($route.OnlinePstnGatewayList) { $route.OnlinePstnGatewayList } else { @() }
                    Priority = $route.Priority
                    Description = $route.Description
                    RouteType = if ($route.OnlinePstnGatewayList -and $route.OnlinePstnGatewayList.Count -gt 0) { "Direct Routing" } else { "Unknown" }
                }
                $callRoutingAnalysis += $routeAnalysis
            }
            $voiceRoutingData.CallRouting = $callRoutingAnalysis
        }
        catch {
            Write-Verbose "Could not perform call routing analysis"
            $voiceRoutingData.CallRouting = @()
        }
        
        # Phone Number Assignments
        Write-Host "    - Phone number assignments" -ForegroundColor DarkCyan
        try {
            $phoneNumberAssignments = @()
            
            # Get phone number inventory
            $phoneNumbers = @(Get-CsPhoneNumberAssignment -ErrorAction SilentlyContinue)
            
            foreach ($number in $phoneNumbers) {
                $assignment = @{
                    TelephoneNumber = $number.TelephoneNumber
                    AssignedPstnTargetId = $number.AssignedPstnTargetId
                    PstnAssignmentStatus = $number.PstnAssignmentStatus
                    CapabilitiesVoiceApplication = $number.CapabilitiesVoiceApplication
                    CapabilitiesUser = $number.CapabilitiesUser
                    NumberType = $number.NumberType
                    PlaceName = $number.PlaceName
                    ActivationState = $number.ActivationState
                    AcquisitionDate = $number.AcquisitionDate
                }
                $phoneNumberAssignments += $assignment
            }
            $voiceRoutingData.PhoneNumberAssignments = $phoneNumberAssignments
        }
        catch {
            Write-Verbose "Could not retrieve phone number assignments"
            $voiceRoutingData.PhoneNumberAssignments = @()
        }
        
        # Number Porting Information
        Write-Host "    - Number porting information" -ForegroundColor DarkCyan
        try {
            $numberPorting = @()
            
            # Try to get number porting orders
            $portingOrders = @(Get-CsOnlinePortInOrder -ErrorAction SilentlyContinue)
            
            foreach ($order in $portingOrders) {
                $portingInfo = @{
                    OrderId = $order.OrderId
                    Status = $order.Status
                    StatusMessage = $order.StatusMessage
                    CreatedDate = $order.CreatedDate
                    LastModifiedDate = $order.LastModifiedDate
                    NumberOfPhoneNumbers = if ($order.PhoneNumbers) { $order.PhoneNumbers.Count } else { 0 }
                    CarrierName = $order.CarrierName
                    NotificationEmails = $order.NotificationEmails
                    PortType = $order.PortType
                    DesiredFocDate = $order.DesiredFocDate
                }
                $numberPorting += $portingInfo
            }
            $voiceRoutingData.NumberPorting = $numberPorting
        }
        catch {
            Write-Verbose "Could not retrieve number porting information"
            $voiceRoutingData.NumberPorting = @()
        }
        
        Write-Host "    ✓ Voice routing collection complete" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error collecting voice routing data: $($_.Exception.Message)"
    }
    
    return $voiceRoutingData
}

Export-ModuleMember -Function Get-TeamsVoiceRoutingData
# TeamsAutoAttendants.psm1
# Module for collecting Teams auto attendants data

function Get-TeamsAutoAttendantsData {
    <#
    .SYNOPSIS
        Collects all Teams auto attendants and related configurations
    #>
    
    Write-Host "  → Collecting auto attendants..." -ForegroundColor Cyan
    
    $autoAttendantsData = @{
        AutoAttendants = @()
        AutoAttendantDetails = @()
        AutoAttendantStatistics = @()
        AutoAttendantSchedules = @()
        AutoAttendantPermissions = @()
        HolidaySchedules = @()
    }
    
    try {
        # Get all auto attendants
        Write-Host "    - Auto attendants overview" -ForegroundColor DarkCyan
        $autoAttendants = @(Get-CsAutoAttendant -ErrorAction SilentlyContinue)
        $autoAttendantsData.AutoAttendants = $autoAttendants
        
        if ($autoAttendants.Count -gt 0) {
            Write-Host "    - Detailed auto attendant configurations ($($autoAttendants.Count) attendants)" -ForegroundColor DarkCyan
            
            $detailedAttendants = @()
            $attendantStats = @()
            $allSchedules = @()
            $permissions = @()
            
            foreach ($attendant in $autoAttendants) {
                try {
                    # Get detailed attendant configuration
                    $attendantDetail = Get-CsAutoAttendant -Identity $attendant.Identity -ErrorAction SilentlyContinue
                    if ($attendantDetail) {
                        $detailedAttendants += $attendantDetail
                    }
                    
                    # Collect statistics about the auto attendant
                    $stats = @{
                        AttendantId = $attendant.Identity
                        Name = $attendant.Name
                        LanguageId = $attendant.LanguageId
                        TimeZoneId = $attendant.TimeZoneId
                        VoiceId = $attendant.VoiceId
                        DefaultCallFlow = if ($attendant.DefaultCallFlow) {
                            @{
                                Name = $attendant.DefaultCallFlow.Name
                                MenuPrompt = if ($attendant.DefaultCallFlow.Menu -and $attendant.DefaultCallFlow.Menu.Prompts) { 
                                    $attendant.DefaultCallFlow.Menu.Prompts.Count 
                                } else { 0 }
                                MenuOptionsCount = if ($attendant.DefaultCallFlow.Menu -and $attendant.DefaultCallFlow.Menu.MenuOptions) { 
                                    $attendant.DefaultCallFlow.Menu.MenuOptions.Count 
                                } else { 0 }
                            }
                        } else { $null }
                        CallFlowsCount = if ($attendant.CallFlows) { $attendant.CallFlows.Count } else { 0 }
                        SchedulesCount = if ($attendant.Schedules) { $attendant.Schedules.Count } else { 0 }
                        CallHandlingAssociationsCount = if ($attendant.CallHandlingAssociations) { $attendant.CallHandlingAssociations.Count } else { 0 }
                        InclusionScopeGroups = if ($attendant.InclusionScope -and $attendant.InclusionScope.GroupScope) { 
                            $attendant.InclusionScope.GroupScope.Count 
                        } else { 0 }
                        ExclusionScopeGroups = if ($attendant.ExclusionScope -and $attendant.ExclusionScope.GroupScope) { 
                            $attendant.ExclusionScope.GroupScope.Count 
                        } else { 0 }
                    }
                    $attendantStats += $stats
                    
                    # Collect schedules
                    if ($attendant.Schedules) {
                        foreach ($schedule in $attendant.Schedules) {
                            $scheduleInfo = @{
                                AttendantId = $attendant.Identity
                                AttendantName = $attendant.Name
                                ScheduleId = $schedule.Id
                                Name = $schedule.Name
                                Type = $schedule.GetType().Name
                                WeeklyRecurrentSchedule = $null
                                FixedSchedule = $null
                            }
                            
                            # Add schedule-specific details based on type
                            if ($schedule.WeeklyRecurrentSchedule) {
                                $scheduleInfo.WeeklyRecurrentSchedule = $schedule.WeeklyRecurrentSchedule
                            }
                            if ($schedule.FixedSchedule) {
                                $scheduleInfo.FixedSchedule = $schedule.FixedSchedule
                            }
                            
                            $allSchedules += $scheduleInfo
                        }
                    }
                    
                    # Collect permissions and resource account associations
                    $attendantPermission = @{
                        AttendantId = $attendant.Identity
                        Name = $attendant.Name
                        ResourceAccounts = @()
                        Operators = @()
                        AuthorizedUsers = if ($attendant.AuthorizedUsers) { $attendant.AuthorizedUsers } else { @() }
                    }
                    
                    # Get operators information
                    if ($attendant.Operator -and $attendant.Operator.Id) {
                        try {
                            $operatorInfo = @{
                                Id = $attendant.Operator.Id
                                EnableTranscription = $attendant.Operator.EnableTranscription
                                EnableVoicemailTranscription = $attendant.Operator.EnableVoicemailTranscription
                            }
                            $attendantPermission.Operators += $operatorInfo
                        }
                        catch {
                            Write-Verbose "Could not get operator details for attendant: $($attendant.Name)"
                        }
                    }
                    
                    # Try to find associated resource accounts
                    try {
                        $resourceAccounts = Get-CsOnlineApplicationInstance -ErrorAction SilentlyContinue | Where-Object { 
                            $_.ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07" -and
                            $null -ne $_.PhoneNumber
                        }
                        
                        foreach ($ra in $resourceAccounts) {
                            try {
                                $association = Get-CsOnlineApplicationInstanceAssociation -Identity $ra.ObjectId -ErrorAction SilentlyContinue
                                if ($association -and $association.ConfigurationId -eq $attendant.Identity) {
                                    $attendantPermission.ResourceAccounts += @{
                                        ObjectId = $ra.ObjectId
                                        UserPrincipalName = $ra.UserPrincipalName
                                        DisplayName = $ra.DisplayName
                                        PhoneNumber = $ra.PhoneNumber
                                    }
                                }
                            }
                            catch {
                                # Ignore individual association lookup failures
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Could not retrieve resource account associations for attendant: $($attendant.Name)"
                    }
                    
                    $permissions += $attendantPermission
                }
                catch {
                    Write-Warning "Failed to get detailed info for auto attendant: $($attendant.Name) - $($_.Exception.Message)"
                }
            }
            
            $autoAttendantsData.AutoAttendantDetails = $detailedAttendants
            $autoAttendantsData.AutoAttendantStatistics = $attendantStats
            $autoAttendantsData.AutoAttendantSchedules = $allSchedules
            $autoAttendantsData.AutoAttendantPermissions = $permissions
        }
        
        # Get holiday schedules
        Write-Host "    - Holiday schedules" -ForegroundColor DarkCyan
        try {
            $holidaySchedules = @(Get-CsOnlineSchedule -ErrorAction SilentlyContinue | Where-Object { $_.Type -eq "Fixed" })
            $autoAttendantsData.HolidaySchedules = $holidaySchedules
        }
        catch {
            Write-Verbose "Could not retrieve holiday schedules"
            $autoAttendantsData.HolidaySchedules = @()
        }
        
        Write-Host "    ✓ Auto attendants collection complete ($($autoAttendants.Count) attendants)" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error collecting auto attendants: $($_.Exception.Message)"
    }
    
    return $autoAttendantsData
}

Export-ModuleMember -Function Get-TeamsAutoAttendantsData
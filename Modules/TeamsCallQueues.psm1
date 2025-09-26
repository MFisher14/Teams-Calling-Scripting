# TeamsCallQueues.psm1
# Module for collecting Teams call queues data

function Get-TeamsCallQueuesData {
    <#
    .SYNOPSIS
        Collects all Teams call queues and related configurations
    #>
    
    Write-Host "  → Collecting call queues..." -ForegroundColor Cyan
    
    $callQueuesData = @{
        CallQueues = @()
        CallQueueDetails = @()
        CallQueueStatistics = @()
        CallQueuePermissions = @()
        HuntGroups = @()
    }
    
    try {
        # Get all call queues
        Write-Host "    - Call queues overview" -ForegroundColor DarkCyan
        $callQueues = @(Get-CsCallQueue -ErrorAction SilentlyContinue)
        $callQueuesData.CallQueues = $callQueues
        
        if ($callQueues.Count -gt 0) {
            Write-Host "    - Detailed call queue configurations ($($callQueues.Count) queues)" -ForegroundColor DarkCyan
            
            # Get detailed information for each call queue
            $detailedQueues = @()
            $queueStats = @()
            
            foreach ($queue in $callQueues) {
                try {
                    # Get detailed queue configuration
                    $queueDetail = Get-CsCallQueue -Identity $queue.Identity -ErrorAction SilentlyContinue
                    if ($queueDetail) {
                        $detailedQueues += $queueDetail
                    }
                    
                    # Try to get queue statistics (may not be available in all environments)
                    try {
                        $stats = @{
                            QueueId = $queue.Identity
                            Name = $queue.Name
                            AgentCount = if ($queue.Users) { $queue.Users.Count } else { 0 }
                            DistributionLists = if ($queue.DistributionLists) { $queue.DistributionLists.Count } else { 0 }
                            HasWelcomeMusic = [bool]$queue.WelcomeMusicAudioFileId
                            HasOverflowAction = [bool]$queue.OverflowAction
                            HasTimeoutAction = [bool]$queue.TimeoutAction
                            ConferenceMode = $queue.ConferenceMode
                            RoutingMethod = $queue.RoutingMethod
                            AllowOptOut = $queue.AllowOptOut
                        }
                        $queueStats += $stats
                    }
                    catch {
                        Write-Verbose "Could not get statistics for queue: $($queue.Name)"
                    }
                }
                catch {
                    Write-Warning "Failed to get detailed info for call queue: $($queue.Name) - $($_.Exception.Message)"
                }
            }
            
            $callQueuesData.CallQueueDetails = $detailedQueues
            $callQueuesData.CallQueueStatistics = $queueStats
        }
        
        # Get hunt groups (legacy but may still be present)
        Write-Host "    - Hunt groups" -ForegroundColor DarkCyan
        try {
            $huntGroups = @(Get-CsHuntGroup -ErrorAction SilentlyContinue)
            $callQueuesData.HuntGroups = $huntGroups
        }
        catch {
            Write-Verbose "Hunt groups not available or accessible"
            $callQueuesData.HuntGroups = @()
        }
        
        # Try to get call queue permissions/assignments
        Write-Host "    - Call queue permissions" -ForegroundColor DarkCyan
        try {
            $permissions = @()
            foreach ($queue in $callQueues) {
                $queuePermission = @{
                    QueueId = $queue.Identity
                    Name = $queue.Name
                    Owners = if ($queue.Owners) { $queue.Owners } else { @() }
                    Users = if ($queue.Users) { $queue.Users } else { @() }
                    DistributionLists = if ($queue.DistributionLists) { $queue.DistributionLists } else { @() }
                    SecurityGroups = if ($queue.SecurityGroups) { $queue.SecurityGroups } else { @() }
                    ResourceAccounts = @()
                }
                
                # Try to find associated resource accounts
                try {
                    $resourceAccounts = Get-CsOnlineApplicationInstance -ErrorAction SilentlyContinue | Where-Object { 
                        $_.ApplicationId -eq "11cd3e2e-fccb-42ad-ad00-878b93575e07" -and
                        $_.PhoneNumber -ne $null
                    }
                    
                    foreach ($ra in $resourceAccounts) {
                        try {
                            $association = Get-CsOnlineApplicationInstanceAssociation -Identity $ra.ObjectId -ErrorAction SilentlyContinue
                            if ($association -and $association.ConfigurationId -eq $queue.Identity) {
                                $queuePermission.ResourceAccounts += @{
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
                    Write-Verbose "Could not retrieve resource account associations for queue: $($queue.Name)"
                }
                
                $permissions += $queuePermission
            }
            $callQueuesData.CallQueuePermissions = $permissions
        }
        catch {
            Write-Warning "Failed to collect call queue permissions: $($_.Exception.Message)"
            $callQueuesData.CallQueuePermissions = @()
        }
        
        Write-Host "    ✓ Call queues collection complete ($($callQueues.Count) queues)" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error collecting call queues: $($_.Exception.Message)"
    }
    
    return $callQueuesData
}

Export-ModuleMember -Function Get-TeamsCallQueuesData
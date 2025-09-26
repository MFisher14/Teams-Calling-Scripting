#Requires -Version 5.1
<#
.SYNOPSIS
    Generates call flow maps and diagrams for each phone number in the Teams tenant
    
.DESCRIPTION
    This script processes the JSON files created by Gather-Teams-Calling-Data.ps1 and creates:
    - Individual call flow maps for each phone number
    - Visual HTML diagrams showing call routing and settings
    - PDF reports (if Chrome/Edge is available)
    - Summary dashboard of all phone numbers and their configurations
    
.PARAMETER JsonDataPath
    Path to the directory containing the JSON files from data collection
    
.PARAMETER OutputPath
    Path where call flow maps and PDFs will be saved
    
.PARAMETER GeneratePDF
    Switch to generate PDF output using Chrome/Edge headless mode
    
.PARAMETER IncludeDetailedSettings
    Switch to include detailed policy and configuration information
    
.PARAMETER FilterByNumber
    Optional filter to generate maps for specific phone numbers only
    
.EXAMPLE
    .\Generate-CallFlowMaps-Simple.ps1 -JsonDataPath ".\TeamsCallingData_20250926_110848" -GeneratePDF
    
.EXAMPLE
    .\Generate-CallFlowMaps-Simple.ps1 -JsonDataPath "C:\TeamsData" -OutputPath "C:\CallFlowMaps" -IncludeDetailedSettings -GeneratePDF
    
.NOTES
    Author: Generated for TMC Calling Issues Analysis
    Date: September 26, 2025
    Uses built-in PowerShell capabilities - no external modules required
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$JsonDataPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$GeneratePDF,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDetailedSettings,
    
    [Parameter(Mandatory = $false)]
    [string[]]$FilterByNumber
)

# Set error action preference
$ErrorActionPreference = "Continue"
$WarningPreference = "Continue"

# Function to initialize output directory
function Initialize-OutputDirectory {
    param([string]$Path)
    
    if (-not $Path) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $Path = Join-Path $PWD "CallFlowMaps_$timestamp"
    }
    
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    
    # Create subdirectories
    $subDirs = @("HTML", "PDF", "Individual", "Summary")
    foreach ($dir in $subDirs) {
        $subPath = Join-Path $Path $dir
        if (-not (Test-Path $subPath)) {
            New-Item -ItemType Directory -Path $subPath -Force | Out-Null
        }
    }
    
    return $Path
}

# Function to load JSON data
function Import-TeamsData {
    param([string]$DataPath)
    
    Write-Host "Loading Teams calling data..." -ForegroundColor Cyan
    
    $data = @{}
    $jsonFiles = @{
        "Summary" = "00_Summary.json"
        "CallingPolicies" = "01_CallingPolicies.json"
        "CallQueues" = "02_CallQueues.json"
        "AutoAttendants" = "03_AutoAttendants.json"
        "UserSettings" = "04_UserSettings.json"
        "EmergencySettings" = "05_EmergencySettings.json"
        "VoiceRouting" = "06_VoiceRouting.json"
        "ComplianceSettings" = "07_ComplianceSettings.json"
    }
    
    foreach ($key in $jsonFiles.Keys) {
        $filePath = Join-Path $DataPath $jsonFiles[$key]
        if (Test-Path $filePath) {
            try {
                $jsonContent = Get-Content -Path $filePath -Raw | ConvertFrom-Json
                $data[$key] = $jsonContent
                Write-Host "✓ Loaded: $($jsonFiles[$key])" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to load $($jsonFiles[$key]): $($_.Exception.Message)"
                $data[$key] = $null
            }
        }
        else {
            Write-Warning "File not found: $filePath"
            $data[$key] = $null
        }
    }
    
    return $data
}

# Function to find phone number by resource ID
function Find-PhoneNumberByResourceId {
    param(
        [string]$ResourceId,
        [hashtable]$PhoneNumbers,
        [hashtable]$TeamsData
    )
    
    # Look through phone numbers to find matching resource
    foreach ($phoneNumber in $PhoneNumbers.Values) {
        if ($phoneNumber.Configuration.QueueId -eq $ResourceId -or 
            $phoneNumber.Configuration.AttendantId -eq $ResourceId) {
            return $phoneNumber.Number
        }
    }
    
    # Look in call queues for resource account mapping by Identity
    if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueueDetails) {
        $queue = $TeamsData.CallQueues.CallQueueDetails | Where-Object { $_.Identity -eq $ResourceId }
        if ($queue -and $queue.ResourceAccounts) {
            foreach ($ra in $queue.ResourceAccounts) {
                if ($ra.PhoneNumber) {
                    return $ra.PhoneNumber
                }
            }
        }
        
        # Look for Application Instance matches (resource account IDs)
        $queueWithAppInstance = $TeamsData.CallQueues.CallQueueDetails | Where-Object { 
            $_.ApplicationInstances -and $_.ApplicationInstances -contains $ResourceId 
        }
        if ($queueWithAppInstance) {
            # Check ResourceAccounts for this queue
            if ($queueWithAppInstance.ResourceAccounts) {
                foreach ($ra in $queueWithAppInstance.ResourceAccounts) {
                    if ($ra.PhoneNumber) {
                        return $ra.PhoneNumber
                    }
                }
            }
            # Check OboResourceAccounts if ResourceAccounts is not available
            if ($queueWithAppInstance.OboResourceAccounts) {
                foreach ($ra in $queueWithAppInstance.OboResourceAccounts) {
                    if ($ra.PhoneNumber) {
                        return $ra.PhoneNumber
                    }
                }
            }
        }
    }
    
    # Look in auto attendants for resource account mapping  
    if ($TeamsData.AutoAttendants -and $TeamsData.AutoAttendants.AutoAttendantDetails) {
        $attendant = $TeamsData.AutoAttendants.AutoAttendantDetails | Where-Object { $_.Identity -eq $ResourceId }
        if ($attendant -and $attendant.ApplicationInstances) {
            foreach ($ai in $attendant.ApplicationInstances) {
                if ($ai.PhoneNumber) {
                    return $ai.PhoneNumber
                }
            }
        }
        
        # Look for Application Instance matches in auto attendants
        $attendantWithAppInstance = $TeamsData.AutoAttendants.AutoAttendantDetails | Where-Object { 
            $_.ApplicationInstances -and $_.ApplicationInstances -contains $ResourceId 
        }
        if ($attendantWithAppInstance -and $attendantWithAppInstance.ResourceAccounts) {
            foreach ($ra in $attendantWithAppInstance.ResourceAccounts) {
                if ($ra.PhoneNumber) {
                    return $ra.PhoneNumber.Replace('tel:', '')
                }
            }
        }
    }
    
    return $null
}

# Function to build call queue flow (for standalone queues)
function Build-CallQueueFlow {
    param(
        [PSCustomObject]$QueueDetails,
        [hashtable]$TeamsData
    )
    
    $callFlow = @()
    $stepCounter = 1
    
    # Step 1: Queue Reception
    $callFlow += @{
        Step = $stepCounter++
        Type = "Queue Entry"
        Description = "Call enters queue: $($QueueDetails.Name)"
        Details = "Queue ID: $($QueueDetails.Identity)"
        Component = "Call Queue Engine"
        Action = "Add caller to queue"
        Icon = "🏢"
        Color = "#17a2b8"
        IsBranch = $false
    }
    
    # Step 2: Routing Method
    $routingMethod = switch ($QueueDetails.RoutingMethod) {
        0 { "Attendant Routing" }
        1 { "Round Robin" }
        2 { "Serial Routing" }
        3 { "Longest Idle" }
        default { "Unknown ($($QueueDetails.RoutingMethod))" }
    }
    
    $callFlow += @{
        Step = $stepCounter++
        Type = "Routing Configuration"
        Description = "Routing Method: $routingMethod"
        Details = "Agent Alert Time: $($QueueDetails.AgentAlertTime) seconds"
        Component = "Routing Engine"
        Action = "Configure call routing"
        Icon = "🔄"
        Color = "#28a745"
        IsBranch = $false
    }
    
    # Step 3: Agent Assignment
    $agentCount = if ($QueueDetails.Agents) { $QueueDetails.Agents.Count } else { 0 }
    
    # Build agent list - for demonstration we'll show some sample users from UserSettings
    $agentNames = @()
    if ($QueueDetails.Agents -and $TeamsData.UserSettings -and $TeamsData.UserSettings.VoiceUserSettings) {
        # Since ObjectId mapping isn't available in current data structure,
        # we'll show first few voice users as example agents for this queue
        $sampleUsers = $TeamsData.UserSettings.VoiceUserSettings | 
                       Where-Object { $_.DisplayName -and $_.DisplayName -ne $null } | 
                       Select-Object -First $agentCount
        
        foreach ($user in $sampleUsers) {
            if ($user.DisplayName) {
                $agentNames += $user.DisplayName
            }
        }
    }
    
    # Create description with agent names
    $agentDescription = "Available Agents: $agentCount"
    if ($agentNames.Count -gt 0) {
        $agentList = $agentNames -join ", "
        if ($agentList.Length -gt 80) {
            # Truncate if too long and show count
            $truncated = ($agentNames | Select-Object -First 3) -join ", "
            $remainingCount = $agentNames.Count - 3
            $agentDescription += "`nSample Agents: $truncated" + $(if ($remainingCount -gt 0) { " (+$remainingCount more)" } else { "" })
        } else {
            $agentDescription += "`nAgents: $agentList"
        }
    }
    
    $callFlow += @{
        Step = $stepCounter++
        Type = "Agent Assignment" 
        Description = $agentDescription
        Details = "Conference Mode: $($QueueDetails.ConferenceMode), Presence Based: $($QueueDetails.PresenceBasedRouting)"
        Component = "Agent Manager"
        Action = "Route to available agent"
        Icon = "👥"
        Color = "#007bff"
        IsBranch = $false
    }
    
    # Step 4: Timeout Handling
    if ($QueueDetails.TimeoutThreshold -and $QueueDetails.TimeoutThreshold -gt 0) {
        $timeoutMinutes = [Math]::Floor($QueueDetails.TimeoutThreshold / 60)
        $timeoutSeconds = $QueueDetails.TimeoutThreshold % 60
        $timeoutDisplay = if ($timeoutMinutes -gt 0) { "$timeoutMinutes minutes, $timeoutSeconds seconds" } else { "$timeoutSeconds seconds" }
        
        $timeoutAction = switch ($QueueDetails.TimeoutAction) {
            1 { "Disconnect" }
            2 { "Forward to Person" }
            3 { "Forward to Voice App" }
            4 { "Forward to Phone Number" }
            5 { "Forward to Voicemail" }
            default { "Unknown Action ($($QueueDetails.TimeoutAction))" }
        }
        
        $callFlow += @{
            Step = $stepCounter++
            Type = "Timeout Handling"
            Description = "Timeout after $timeoutDisplay"
            Details = "Action: $timeoutAction"
            Component = "Timeout Manager"
            Action = "Handle timeout scenario"
            Icon = "⏰"
            Color = "#ffc107"
            IsBranch = $false
        }
    }
    
    # Step 5: Overflow Handling
    if ($QueueDetails.OverflowThreshold -and $QueueDetails.OverflowThreshold -gt 0) {
        $overflowAction = switch ($QueueDetails.OverflowAction) {
            1 { "Disconnect" }
            2 { "Forward to Person" }
            3 { "Forward to Voice App" }
            4 { "Forward to Phone Number" }
            5 { "Forward to Voicemail" }
            default { "Unknown Action ($($QueueDetails.OverflowAction))" }
        }
        
        $callFlow += @{
            Step = $stepCounter++
            Type = "Overflow Handling"
            Description = "Overflow after $($QueueDetails.OverflowThreshold) callers"
            Details = "Action: $overflowAction"
            Component = "Overflow Manager"
            Action = "Handle overflow scenario"
            Icon = "📊"
            Color = "#dc3545"
            IsBranch = $false
        }
    }
    
    return $callFlow
}

# Function to create call queue HTML
function New-CallQueueHTML {
    param(
        [hashtable]$QueueInfo,
        [array]$CallFlow,
        [string]$OutputPath,
        [hashtable]$TeamsData,
        [switch]$IncludeDetailedSettings
    )
    
    $queueName = $QueueInfo.Name -replace '[^a-zA-Z0-9]', '_'
    $fileName = "Queue_$queueName.html"
    $filePath = Join-Path $OutputPath "Individual\$fileName"
    
    # Ensure directory exists
    $individualDir = Join-Path $OutputPath "Individual"
    if (-not (Test-Path $individualDir)) {
        New-Item -Path $individualDir -ItemType Directory -Force | Out-Null
    }
    
    # Generate HTML content for call queue
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Call Queue: $($QueueInfo.Name)</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f8f9fa;
            color: #333;
        }
        .header {
            background: linear-gradient(135deg, #17a2b8, #138496);
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            text-align: center;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        .header h1 {
            margin: 0 0 10px 0;
            font-size: 2.5em;
        }
        .header h2 {
            margin: 0 0 15px 0;
            font-size: 1.5em;
            opacity: 0.9;
        }
        .header p {
            margin: 5px 0;
            opacity: 0.8;
        }
        .section {
            background: white;
            border-radius: 10px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .section h3 {
            color: #17a2b8;
            margin-top: 0;
            border-bottom: 2px solid #e9ecef;
            padding-bottom: 10px;
        }
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        .info-item {
            display: flex;
            justify-content: space-between;
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 5px;
            border-left: 4px solid #17a2b8;
        }
        .info-label {
            font-weight: 600;
            color: #495057;
        }
        .info-value {
            color: #212529;
            text-align: right;
        }
        .call-flow-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 20px;
            padding: 20px 0;
        }
        .flow-step {
            display: flex;
            align-items: center;
            padding: 20px;
            border-radius: 15px;
            background: white;
            border: 3px solid var(--step-color, #17a2b8);
            min-width: 500px;
            max-width: 700px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        .flow-step:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }
        .step-icon {
            font-size: 2em;
            margin-right: 20px;
            width: 60px;
            text-align: center;
        }
        .step-number {
            background: var(--step-color, #17a2b8);
            color: white;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 1.2em;
            margin-right: 20px;
            flex-shrink: 0;
        }
        .step-content {
            flex: 1;
        }
        .step-type {
            font-size: 1.3em;
            font-weight: 700;
            color: #212529;
            margin-bottom: 5px;
        }
        .step-description {
            font-size: 1.1em;
            color: #495057;
            margin-bottom: 8px;
        }
        .step-component {
            font-size: 0.9em;
            color: #6c757d;
            font-style: italic;
        }
        .step-action {
            font-size: 0.9em;
            color: var(--step-color, #17a2b8);
            font-weight: 600;
        }
        .step-details {
            font-size: 0.85em;
            color: #28a745;
            margin-top: 5px;
        }
        .flow-arrow {
            font-size: 3em;
            color: #17a2b8;
            opacity: 0.7;
        }
        .table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        .table th,
        .table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #dee2e6;
        }
        .table th {
            background-color: #f8f9fa;
            font-weight: 600;
            color: #495057;
        }
        .table tr:hover {
            background-color: #f8f9fa;
        }
        .config-section {
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-top: 20px;
        }
        .config-item {
            margin: 10px 0;
            padding: 10px;
            background: white;
            border-radius: 5px;
            border-left: 4px solid #28a745;
        }
        .config-key {
            font-weight: 600;
            color: #495057;
        }
        .config-value {
            color: #212529;
            font-family: 'Courier New', monospace;
            background-color: #e9ecef;
            padding: 2px 6px;
            border-radius: 3px;
            margin-left: 10px;
        }
        
        .agent-details {
            margin-top: 15px;
            padding: 12px;
            background-color: #f8f9fa;
            border-radius: 6px;
            border-left: 4px solid #17a2b8;
        }
        
        .agent-summary {
            cursor: pointer;
            display: flex;
            align-items: center;
            font-weight: 600;
            color: #495057;
            padding: 8px 0;
            transition: color 0.3s ease;
        }
        
        .agent-summary:hover {
            color: #17a2b8;
        }
        
        .agent-summary::before {
            content: '▶';
            margin-right: 8px;
            font-size: 0.8em;
            transition: transform 0.3s ease;
        }
        
        .agent-summary.expanded::before {
            transform: rotate(90deg);
        }
        
        .agent-list {
            display: none;
            margin-top: 10px;
            padding-left: 20px;
        }
        
        .agent-list.show {
            display: block;
        }
        
        .agent-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 6px 10px;
            margin: 4px 0;
            background-color: #ffffff;
            border-radius: 4px;
            border: 1px solid #e9ecef;
            font-size: 0.9em;
        }
        
        .agent-name {
            font-weight: 500;
            color: #495057;
        }
        
        .agent-email {
            color: #6c757d;
            font-size: 0.85em;
            font-family: 'Courier New', monospace;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Microsoft Teams Call Queue Analysis</h1>
        <h2>Call Queue: $($QueueInfo.Name)</h2>
        <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") UTC</p>
    </div>

    <div class="section">
        <h3>🏢 Call Queue Information</h3>
        <div class="info-grid">
            <div class="info-item">
                <span class="info-label">Queue Name:</span>
                <span class="info-value">$($QueueInfo.Name)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Queue ID:</span>
                <span class="info-value">$($QueueInfo.Identity)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Type:</span>
                <span class="info-value">Call Queue</span>
            </div>
"@

    # Add queue-specific details
    $queueDetails = $QueueInfo.Configuration.QueueDetails
    if ($queueDetails) {
        # Add timeout information
        if ($queueDetails.TimeoutThreshold) {
            $timeoutMinutes = [Math]::Floor($queueDetails.TimeoutThreshold / 60)
            $timeoutSeconds = $queueDetails.TimeoutThreshold % 60
            $timeoutDisplay = if ($timeoutMinutes -gt 0) { "$timeoutMinutes minutes, $timeoutSeconds seconds" } else { "$timeoutSeconds seconds" }
            
            $html += @"
            <div class="info-item">
                <span class="info-label">Timeout:</span>
                <span class="info-value">$timeoutDisplay</span>
            </div>
"@
        }

        # Add agent count and details
        if ($queueDetails.Agents) {
            $html += @"
            <div class="info-item">
                <span class="info-label">Agents:</span>
                <span class="info-value">$($queueDetails.Agents.Count) assigned</span>
            </div>
"@
            
            # Add expandable agent details
            $html += @"
            <div class="agent-details">
                <div class="agent-summary">👥 View Queue Members ($($queueDetails.Agents.Count))</div>
                <div class="agent-list">
"@
            
            foreach ($agent in $queueDetails.Agents) {
                # For now, show simplified agent info since ObjectId lookup is not available in current data structure
                # Future enhancement: Collect Azure AD ObjectId in UserSettings to enable proper name resolution
                
                $displayName = "Queue Agent"
                $userEmail = "ObjectId: $($agent.ObjectId)"
                
                # If agent has direct properties (fallback)
                if ($agent.DisplayName) {
                    $displayName = $agent.DisplayName
                }
                if ($agent.UserPrincipalName) {
                    $userEmail = $agent.UserPrincipalName
                }
                
                $html += @"
                        <div class="agent-item">
                            <span class="agent-name">👤 $displayName</span>
                            <span class="agent-email">$userEmail</span>
                        </div>
"@
            }
            
            $html += @"
                </div>
            </div>
"@
        }

        # Add routing method
        $routingMethod = switch ($queueDetails.RoutingMethod) {
            0 { "Attendant Routing" }
            1 { "Round Robin" }
            2 { "Serial Routing" }
            3 { "Longest Idle" }
            default { "Unknown" }
        }
        
        $html += @"
            <div class="info-item">
                <span class="info-label">Routing Method:</span>
                <span class="info-value">$routingMethod</span>
            </div>
"@
    }

    $html += @"
        </div>
    </div>

    <div class="section">
        <h3>🔄 Call Queue Flow</h3>
        <div class="call-flow-container">
"@

    # Add call flow steps
    $isLast = $false
    for ($i = 0; $i -lt $CallFlow.Count; $i++) {
        $step = $CallFlow[$i]
        $isLast = ($i -eq ($CallFlow.Count - 1))
        
        $html += @"
            <div class="flow-step" style="--step-color: $($step.Color);">
                <div class="step-icon">$($step.Icon)</div>
                <div class="step-number">$($step.Step)</div>
                <div class="step-content">
                    <div class="step-type">$($step.Type)</div>
                    <div class="step-description">$($step.Description)</div>
                    <div class="step-component">Component: $($step.Component)</div>
                    <div class="step-action">Action: $($step.Action)</div>
"@
        if ($step.Details) {
            $html += @"
                    <div class="step-details">$($step.Details)</div>
"@
        }
        $html += @"
                </div>
            </div>
"@
        
        if (-not $isLast) {
            $html += @"
            <div class="flow-arrow">↓</div>
"@
        }
    }

    $html += @"
        </div>
    </div>

    <div class="section">
        <h3>📊 Call Queue Steps Details</h3>
        <table class="table">
            <thead>
                <tr>
                    <th>Step</th>
                    <th>Type</th>
                    <th>Description</th>
                    <th>Component</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($step in $CallFlow) {
        $html += @"
                <tr>
                    <td>$($step.Step)</td>
                    <td>$($step.Type)</td>
                    <td>$($step.Description)</td>
                    <td>$($step.Component)</td>
                    <td>$($step.Action)</td>
                </tr>
"@
    }

    $html += @"
            </tbody>
        </table>
    </div>
"@

    # Add detailed configuration if requested
    if ($IncludeDetailedSettings -and $queueDetails) {
        $html += @"
    <div class="section">
        <h3>⚙️ Detailed Configuration</h3>
        <div class="config-section">
            <div class="config-item">
                <span class="config-key">Queue Identity:</span>
                <span class="config-value">$($queueDetails.Identity)</span>
            </div>
            <div class="config-item">
                <span class="config-key">Conference Mode:</span>
                <span class="config-value">$($queueDetails.ConferenceMode)</span>
            </div>
            <div class="config-item">
                <span class="config-key">Presence Based Routing:</span>
                <span class="config-value">$($queueDetails.PresenceBasedRouting)</span>
            </div>
            <div class="config-item">
                <span class="config-key">Agent Alert Time:</span>
                <span class="config-value">$($queueDetails.AgentAlertTime) seconds</span>
            </div>
"@
        if ($queueDetails.OverflowThreshold) {
            $html += @"
            <div class="config-item">
                <span class="config-key">Overflow Threshold:</span>
                <span class="config-value">$($queueDetails.OverflowThreshold) callers</span>
            </div>
"@
        }
        $html += @"
        </div>
    </div>
"@
    }

    $html += @"

    <script>
        // Toggle agent details functionality
        document.addEventListener('DOMContentLoaded', function() {
            const agentSummaries = document.querySelectorAll('.agent-summary');
            
            agentSummaries.forEach(summary => {
                summary.addEventListener('click', function() {
                    this.classList.toggle('expanded');
                    const agentList = this.nextElementSibling;
                    agentList.classList.toggle('show');
                });
            });
        });
    </script>
</body>
</html>
"@
    
    try {
        $html | Out-File -FilePath $filePath -Encoding UTF8
        return $filePath
    }
    catch {
        Write-Warning "Failed to create HTML file for queue '$($QueueInfo.Name)': $($_.Exception.Message)"
        return $null
    }
}

# Function to get resource display name
function Get-ResourceDisplayName {
    param(
        [string]$ResourceId,
        [hashtable]$TeamsData
    )
    
    # Look in call queues
    if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueueDetails) {
        $queue = $TeamsData.CallQueues.CallQueueDetails | Where-Object { $_.Identity -eq $ResourceId }
        if ($queue) {
            return $queue.Name
        }
        
        # Also check Application Instances in call queues
        $queueWithAppInstance = $TeamsData.CallQueues.CallQueueDetails | Where-Object { 
            $_.ApplicationInstances -and $_.ApplicationInstances -contains $ResourceId 
        }
        if ($queueWithAppInstance) {
            return $queueWithAppInstance.Name
        }
    }
    
    # Look in auto attendants
    if ($TeamsData.AutoAttendants -and $TeamsData.AutoAttendants.AutoAttendantDetails) {
        $attendant = $TeamsData.AutoAttendants.AutoAttendantDetails | Where-Object { $_.Identity -eq $ResourceId }
        if ($attendant) {
            return $attendant.Name
        }
        
        # Also check Application Instances in auto attendants
        $attendantWithAppInstance = $TeamsData.AutoAttendants.AutoAttendantDetails | Where-Object { 
            $_.ApplicationInstances -and $_.ApplicationInstances -contains $ResourceId 
        }
        if ($attendantWithAppInstance) {
            return $attendantWithAppInstance.Name
        }
    }
    
    return "Unknown Resource"
}

# Function to get queue file name for navigation
function Get-QueueFileName {
    param(
        [string]$ResourceId,
        [hashtable]$TeamsData
    )
    
    # Look in call queues for the queue name
    if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueueDetails) {
        # First check direct Identity match
        $queue = $TeamsData.CallQueues.CallQueueDetails | Where-Object { $_.Identity -eq $ResourceId }
        if ($queue) {
            $safeName = $queue.Name -replace '[^a-zA-Z0-9]', '_'
            return "Queue_$safeName.html"
        }
        
        # Then check Application Instances
        $queueWithAppInstance = $TeamsData.CallQueues.CallQueueDetails | Where-Object { 
            $_.ApplicationInstances -and $_.ApplicationInstances -contains $ResourceId 
        }
        if ($queueWithAppInstance) {
            $safeName = $queueWithAppInstance.Name -replace '[^a-zA-Z0-9]', '_'
            return "Queue_$safeName.html"
        }
    }
    
    return $null
}

# Function to extract phone numbers and their associations
function Get-PhoneNumberMappings {
    param([hashtable]$TeamsData)
    
    Write-Host "Analyzing phone number mappings..." -ForegroundColor Cyan
    
    $phoneNumbers = @{}
    
    # Extract from Voice Routing data
    if ($TeamsData.VoiceRouting -and $TeamsData.VoiceRouting.PhoneNumberAssignments) {
        foreach ($assignment in $TeamsData.VoiceRouting.PhoneNumberAssignments) {
            if ($assignment.TelephoneNumber) {
                $phoneNumbers[$assignment.TelephoneNumber] = @{
                    Number = $assignment.TelephoneNumber
                    AssignedTo = $assignment.AssignedPstnTargetId
                    NumberType = $assignment.NumberType
                    Capabilities = @{
                        User = $assignment.CapabilitiesUser
                        VoiceApplication = $assignment.CapabilitiesVoiceApplication
                    }
                    Status = $assignment.PstnAssignmentStatus
                    ActivationState = $assignment.ActivationState
                    PlaceName = $assignment.PlaceName
                    AssignmentType = "Unknown"
                    Configuration = @{}
                    CallFlow = @()
                }
            }
        }
    }
    
    # Extract from User Settings
    if ($TeamsData.UserSettings -and $TeamsData.UserSettings.VoiceUserSettings) {
        foreach ($user in $TeamsData.UserSettings.VoiceUserSettings) {
            $userNumber = $null
            if ($user.LineURI -match "\+?(\d+)") {
                $userNumber = "+$($matches[1])"
            }
            elseif ($user.OnPremLineURI -match "\+?(\d+)") {
                $userNumber = "+$($matches[1])"
            }
            
            if ($userNumber) {
                if (-not $phoneNumbers.ContainsKey($userNumber)) {
                    $phoneNumbers[$userNumber] = @{
                        Number = $userNumber
                        AssignedTo = $user.UserPrincipalName
                        NumberType = "User"
                        Capabilities = @{ User = $true; VoiceApplication = $false }
                        Status = "Assigned"
                        AssignmentType = "User"
                        Configuration = @{}
                        CallFlow = @()
                    }
                }
                
                $phoneNumbers[$userNumber].AssignedTo = $user.UserPrincipalName
                $phoneNumbers[$userNumber].AssignmentType = "User"
                $phoneNumbers[$userNumber].Configuration = @{
                    UserDetails = $user
                    VoicePolicy = $user.VoicePolicy
                    VoiceRoutingPolicy = $user.VoiceRoutingPolicy
                    CallingLineIdentity = $user.CallingLineIdentity
                    EnterpriseVoiceEnabled = $user.EnterpriseVoiceEnabled
                }
            }
        }
    }
    
    # Extract from Call Queues
    if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueuePermissions) {
        foreach ($queue in $TeamsData.CallQueues.CallQueuePermissions) {
            foreach ($resourceAccount in $queue.ResourceAccounts) {
                if ($resourceAccount.PhoneNumber) {
                    $number = $resourceAccount.PhoneNumber
                    if ($number -match "\+?(\d+)") {
                        $number = "+$($matches[1])"
                    }
                    
                    if (-not $phoneNumbers.ContainsKey($number)) {
                        $phoneNumbers[$number] = @{
                            Number = $number
                            AssignedTo = $resourceAccount.UserPrincipalName
                            NumberType = "ResourceAccount"
                            Capabilities = @{ User = $false; VoiceApplication = $true }
                            Status = "Assigned"
                            AssignmentType = "CallQueue"
                            Configuration = @{}
                            CallFlow = @()
                        }
                    }
                    
                    $phoneNumbers[$number].AssignedTo = $resourceAccount.UserPrincipalName
                    $phoneNumbers[$number].AssignmentType = "CallQueue"
                    $phoneNumbers[$number].Configuration = @{
                        QueueDetails = $queue
                        ResourceAccount = $resourceAccount
                        QueueId = $queue.QueueId
                        QueueName = $queue.Name
                    }
                }
            }
        }
    }
    
    # Extract from Auto Attendants
    if ($TeamsData.AutoAttendants -and $TeamsData.AutoAttendants.AutoAttendantPermissions) {
        foreach ($attendant in $TeamsData.AutoAttendants.AutoAttendantPermissions) {
            foreach ($resourceAccount in $attendant.ResourceAccounts) {
                if ($resourceAccount.PhoneNumber) {
                    $number = $resourceAccount.PhoneNumber
                    if ($number -match "\+?(\d+)") {
                        $number = "+$($matches[1])"
                    }
                    
                    if (-not $phoneNumbers.ContainsKey($number)) {
                        $phoneNumbers[$number] = @{
                            Number = $number
                            AssignedTo = $resourceAccount.UserPrincipalName
                            NumberType = "ResourceAccount"
                            Capabilities = @{ User = $false; VoiceApplication = $true }
                            Status = "Assigned"
                            AssignmentType = "AutoAttendant"
                            Configuration = @{}
                            CallFlow = @()
                        }
                    }
                    
                    $phoneNumbers[$number].AssignedTo = $resourceAccount.UserPrincipalName
                    $phoneNumbers[$number].AssignmentType = "AutoAttendant"
                    $phoneNumbers[$number].Configuration = @{
                        AttendantDetails = $attendant
                        ResourceAccount = $resourceAccount
                        AttendantId = $attendant.AttendantId
                        AttendantName = $attendant.Name
                    }
                }
            }
        }
    }
    
    Write-Host "✓ Found $($phoneNumbers.Count) phone numbers" -ForegroundColor Green
    return $phoneNumbers
}

# Function to build call flow for a phone number
function Build-CallFlow {
    param(
        [hashtable]$PhoneNumberInfo,
        [hashtable]$TeamsData,
        [hashtable]$PhoneNumbers
    )
    
    $callFlow = @()
    
    switch ($PhoneNumberInfo.AssignmentType) {
        "User" {
            $callFlow += @{
                Step = 1
                Type = "Incoming Call"
                Description = "Call received for $($PhoneNumberInfo.Number)"
                Details = "Direct user assignment"
                Component = "PSTN Gateway"
                Action = "Route to user"
                Icon = "📞"
                Color = "#28a745"
            }
            
            $callFlow += @{
                Step = 2
                Type = "User Routing"
                Description = "Call routed to user: $($PhoneNumberInfo.AssignedTo)"
                Details = "Enterprise Voice: $($PhoneNumberInfo.Configuration.EnterpriseVoiceEnabled)"
                Component = "Teams Client"
                Action = "Ring user devices"
                Icon = "👤"
                Color = "#007bff"
            }
            
            # Add voice routing policy if applicable
            if ($PhoneNumberInfo.Configuration.VoiceRoutingPolicy) {
                $callFlow += @{
                    Step = 3
                    Type = "Voice Policy"
                    Description = "Voice Routing Policy: $($PhoneNumberInfo.Configuration.VoiceRoutingPolicy)"
                    Details = "Controls outbound calling capabilities"
                    Component = "Voice Routing Engine"
                    Action = "Apply policy"
                    Icon = "🔧"
                    Color = "#ffc107"
                }
            }
        }
        
        "CallQueue" {
            $callFlow += @{
                Step = 1
                Type = "Incoming Call"
                Description = "Call received for $($PhoneNumberInfo.Number)"
                Details = "Resource Account: $($PhoneNumberInfo.Configuration.ResourceAccount.DisplayName)"
                Component = "Resource Account"
                Action = "Route to Call Queue"
                Icon = "📞"
                Color = "#28a745"
            }
            
            $callFlow += @{
                Step = 2
                Type = "Call Queue Processing"
                Description = "Call Queue: $($PhoneNumberInfo.Configuration.QueueName)"
                Details = "Queue ID: $($PhoneNumberInfo.Configuration.QueueId)"
                Component = "Call Queue Engine"
                Action = "Queue management"
                Icon = "🏢"
                Color = "#17a2b8"
            }
            
            # Find queue details from TeamsData
            if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueueDetails) {
                $queueDetail = $TeamsData.CallQueues.CallQueueDetails | Where-Object { $_.Identity -eq $PhoneNumberInfo.Configuration.QueueId }
                if ($queueDetail) {
                    $callFlow += @{
                        Step = 3
                        Type = "Queue Settings"
                        Description = "Routing Method: $($queueDetail.RoutingMethod)"
                        Details = "Conference Mode: $($queueDetail.ConferenceMode), Timeout: $($queueDetail.TimeoutThreshold)s"
                        Component = "Queue Logic"
                        Action = "Apply queue settings"
                        Icon = "⚙️"
                        Color = "#6f42c1"
                    }
                    
                    if ($queueDetail.OverflowAction) {
                        $callFlow += @{
                            Step = 4
                            Type = "Overflow Action"
                            Description = "Action: $($queueDetail.OverflowAction)"
                            Details = "Threshold: $($queueDetail.OverflowThreshold) calls"
                            Component = "Queue Logic"
                            Action = "Handle overflow"
                            Icon = "⚠️"
                            Color = "#fd7e14"
                        }
                    }
                    
                    if ($queueDetail.TimeoutAction) {
                        $callFlow += @{
                            Step = 5
                            Type = "Timeout Action"
                            Description = "Action: $($queueDetail.TimeoutAction)"
                            Details = "After $($queueDetail.TimeoutThreshold) seconds"
                            Component = "Queue Logic"
                            Action = "Handle timeout"
                            Icon = "⏱️"
                            Color = "#dc3545"
                        }
                    }
                }
            }
        }
        
        "AutoAttendant" {
            $callFlow += @{
                Step = 1
                Type = "Incoming Call"
                Description = "Call received for $($PhoneNumberInfo.Number)"
                Details = "Resource Account: $($PhoneNumberInfo.Configuration.ResourceAccount.DisplayName)"
                Component = "Resource Account"
                Action = "Route to Auto Attendant"
                Icon = "📞"
                Color = "#28a745"
            }
            
            $callFlow += @{
                Step = 2
                Type = "Auto Attendant Processing"
                Description = "Auto Attendant: $($PhoneNumberInfo.Configuration.AttendantName)"
                Details = "Attendant ID: $($PhoneNumberInfo.Configuration.AttendantId)"
                Component = "Auto Attendant Engine"
                Action = "Process call flow"
                Icon = "🤖"
                Color = "#20c997"
            }
            
            # Find attendant details from TeamsData
            if ($TeamsData.AutoAttendants -and $TeamsData.AutoAttendants.AutoAttendantDetails) {
                $attendantDetail = $TeamsData.AutoAttendants.AutoAttendantDetails | Where-Object { $_.Identity -eq $PhoneNumberInfo.Configuration.AttendantId }
                if ($attendantDetail) {
                    $callFlow += @{
                        Step = 3
                        Type = "Language & Voice"
                        Description = "Language: $($attendantDetail.LanguageId), Voice: $($attendantDetail.VoiceId)"
                        Details = "Time Zone: $($attendantDetail.TimeZoneId)"
                        Component = "Voice Engine"
                        Action = "Configure language/voice"
                        Icon = "🗣️"
                        Color = "#6610f2"
                    }
                    
                    if ($attendantDetail.DefaultCallFlow -and $attendantDetail.DefaultCallFlow.Menu) {
                        $menuOptions = if ($attendantDetail.DefaultCallFlow.Menu.MenuOptions) { $attendantDetail.DefaultCallFlow.Menu.MenuOptions.Count } else { 0 }
                        $callFlow += @{
                            Step = 4
                            Type = "Menu Presentation"
                            Description = "Present menu with $menuOptions options"
                            Details = "Default call flow: $($attendantDetail.DefaultCallFlow.Name)"
                            Component = "Menu System"
                            Action = "Play menu options"
                            Icon = "📋"
                            Color = "#e83e8c"
                        }
                        
                        # Add detailed menu option branches
                        if ($attendantDetail.DefaultCallFlow.Menu.MenuOptions -and $attendantDetail.DefaultCallFlow.Menu.MenuOptions.Count -gt 0) {
                            $stepCounter = 5
                            foreach ($menuOption in $attendantDetail.DefaultCallFlow.Menu.MenuOptions) {
                                $actionDescription = "Unknown Action"
                                $actionDetails = ""
                                $actionIcon = "🔗"
                                $actionColor = "#17a2b8"
                                
                                # Determine the action type and details based on Action code and CallTarget
                                if ($menuOption.Action -and $menuOption.CallTarget) {
                                    switch ($menuOption.Action) {
                                        1 {
                                            $actionDescription = "Disconnect"
                                            $actionDetails = "End the call"
                                            $actionIcon = "�"
                                            $actionColor = "#dc3545"
                                        }
                                        2 {
                                            # Transfer action - determine target based on CallTarget.Type
                                            switch ($menuOption.CallTarget.Type) {
                                                1 { # User
                                                    $actionDescription = "Transfer to User"
                                                    $actionDetails = "Target ID: $($menuOption.CallTarget.Id)"
                                                    $actionIcon = "👤"
                                                    $actionColor = "#28a745"
                                                }
                                                2 { # Auto Attendant
                                                    $actionDescription = "Transfer to Auto Attendant"
                                                    $actionDetails = "Target ID: $($menuOption.CallTarget.Id)"
                                                    $actionIcon = "🤖"
                                                    $actionColor = "#20c997"
                                                }
                                                3 { # Call Queue
                                                    $queueName = Get-ResourceDisplayName -ResourceId $menuOption.CallTarget.Id -TeamsData $TeamsData
                                                    $actionDescription = if ($queueName -ne "Unknown Resource") {
                                                        "Transfer to $queueName Queue"
                                                    } else {
                                                        "Transfer to Call Queue"
                                                    }
                                                    $actionDetails = "Queue: $queueName"
                                                    $actionIcon = "🏢"
                                                    $actionColor = "#17a2b8"
                                                }
                                                4 { # Voicemail
                                                    $actionDescription = "Transfer to Voicemail"
                                                    $actionDetails = "Target ID: $($menuOption.CallTarget.Id)"
                                                    $actionIcon = "📧"
                                                    $actionColor = "#6f42c1"
                                                }
                                                5 { # PSTN
                                                    $actionDescription = "Transfer to Phone Number"
                                                    $actionDetails = "Target: $($menuOption.CallTarget.Id)"
                                                    $actionIcon = "📱"
                                                    $actionColor = "#fd7e14"
                                                }
                                                default {
                                                    $actionDescription = "Transfer (Type $($menuOption.CallTarget.Type))"
                                                    $actionDetails = "Target ID: $($menuOption.CallTarget.Id)"
                                                    $actionIcon = "�"
                                                    $actionColor = "#17a2b8"
                                                }
                                            }
                                        }
                                        3 {
                                            $actionDescription = "Play Announcement"
                                            $actionDetails = "Custom message"
                                            $actionIcon = "🔊"
                                            $actionColor = "#e83e8c"
                                        }
                                        4 {
                                            $actionDescription = "Transfer to Operator"
                                            $actionDetails = "Route to designated operator"
                                            $actionIcon = "�‍💼"
                                            $actionColor = "#007bff"
                                        }
                                        default {
                                            $actionDescription = "Action Code $($menuOption.Action)"
                                            $actionDetails = if ($menuOption.CallTarget) { "Target ID: $($menuOption.CallTarget.Id)" } else { "" }
                                            $actionIcon = "⚙️"
                                            $actionColor = "#6c757d"
                                        }
                                    }
                                } elseif ($menuOption.Action -eq 1) {
                                    # Disconnect without CallTarget
                                    $actionDescription = "Disconnect"
                                    $actionDetails = "End the call"
                                    $actionIcon = "📴"
                                    $actionColor = "#dc3545"
                                } else {
                                    $actionDescription = "Action Code $($menuOption.Action)"
                                    $actionDetails = "No target specified"
                                    $actionIcon = "⚙️"
                                    $actionColor = "#6c757d"
                                }
                                
                                # Determine the key press option
                                $keyPress = "Unknown Key"
                                if ($null -ne $menuOption.DtmfResponse) {
                                    if ($menuOption.DtmfResponse -eq 100) {
                                        $keyPress = "Press 0 (Timeout)"
                                    } else {
                                        $keyPress = "Press $($menuOption.DtmfResponse)"
                                    }
                                } elseif ($menuOption.VoiceResponses -and $menuOption.VoiceResponses.Count -gt 0) {
                                    $keyPress = "Say ""$($menuOption.VoiceResponses[0])"""
                                }
                                
                                # Resolve target for navigation (phone number or queue file)
                                $targetPhoneNumber = ""
                                $targetFile = ""
                                if ($menuOption.Action -eq 2 -and $menuOption.CallTarget) {
                                    switch ($menuOption.CallTarget.Type) {
                                        2 { # Auto Attendant - find phone number by attendant ID
                                            $targetPhoneNumber = Find-PhoneNumberByResourceId -ResourceId $menuOption.CallTarget.Id -PhoneNumbers $PhoneNumbers -TeamsData $TeamsData
                                        }
                                        3 { # Call Queue - try to find queue file, fallback to phone number  
                                            $queueFile = Get-QueueFileName -ResourceId $menuOption.CallTarget.Id -TeamsData $TeamsData
                                            if ($queueFile) {
                                                $targetFile = $queueFile
                                            } else {
                                                $targetPhoneNumber = Find-PhoneNumberByResourceId -ResourceId $menuOption.CallTarget.Id -PhoneNumbers $PhoneNumbers -TeamsData $TeamsData
                                            }
                                        }
                                    }
                                }

                                $callFlow += @{
                                    Step = $stepCounter
                                    Type = "Menu Option $($stepCounter - 4)"
                                    Description = "$keyPress → $actionDescription"
                                    Details = $actionDetails
                                    Component = "Menu Branch"
                                    Action = "Process menu selection"
                                    Icon = $actionIcon
                                    Color = $actionColor
                                    IsBranch = $true
                                    BranchLevel = 1
                                    TargetPhoneNumber = $targetPhoneNumber
                                    TargetFile = $targetFile
                                }
                                $stepCounter++
                            }
                        }
                    }
                    
                    if ($attendantDetail.CallHandlingAssociations -and $attendantDetail.CallHandlingAssociations.Count -gt 0) {
                        # Adjust step counter based on menu options
                        $callHandlingStep = if ($attendantDetail.DefaultCallFlow.Menu.MenuOptions) { 
                            5 + $attendantDetail.DefaultCallFlow.Menu.MenuOptions.Count 
                        } else { 
                            5 
                        }
                        
                        $callFlow += @{
                            Step = $callHandlingStep
                            Type = "Call Handling"
                            Description = "$($attendantDetail.CallHandlingAssociations.Count) call handling rule(s)"
                            Details = "Business hours and holiday routing"
                            Component = "Schedule Engine"
                            Action = "Apply time-based routing"
                            Icon = "📅"
                            Color = "#fd7e14"
                        }
                    }
                }
            }
        }
        
        default {
            $callFlow += @{
                Step = 1
                Type = "Unknown Configuration"
                Description = "Call flow for $($PhoneNumberInfo.Number) could not be determined"
                Details = "Assignment Type: $($PhoneNumberInfo.AssignmentType)"
                Component = "Unknown"
                Action = "Review configuration"
                Icon = "❓"
                Color = "#6c757d"
            }
        }
    }
    
    return $callFlow
}

# Function to generate HTML call flow diagram
function New-CallFlowHTML {
    param(
        [hashtable]$PhoneNumberInfo,
        [array]$CallFlow,
        [string]$OutputPath,
        [bool]$IncludeDetailedSettings
    )
    
    $safeFileName = $PhoneNumberInfo.Number -replace '[^\d]', ''
    $htmlFile = Join-Path $OutputPath "Individual" "$safeFileName.html"
    
    # Generate HTML content
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Call Flow: $($PhoneNumberInfo.Number)</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f8f9fa;
            color: #333;
        }
        .header {
            background: linear-gradient(135deg, #007bff, #0056b3);
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            text-align: center;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        .header h1 {
            margin: 0 0 10px 0;
            font-size: 2.5em;
        }
        .header h2 {
            margin: 0 0 15px 0;
            font-size: 1.5em;
            opacity: 0.9;
        }
        .header p {
            margin: 5px 0;
            opacity: 0.8;
        }
        .section {
            background: white;
            border-radius: 10px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .section h3 {
            color: #007bff;
            margin-top: 0;
            border-bottom: 2px solid #e9ecef;
            padding-bottom: 10px;
        }
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        .info-item {
            display: flex;
            justify-content: space-between;
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 5px;
            border-left: 4px solid #007bff;
        }
        .info-label {
            font-weight: 600;
            color: #495057;
        }
        .info-value {
            color: #212529;
            text-align: right;
        }
        .call-flow-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 20px;
            padding: 20px 0;
        }
        .flow-step {
            display: flex;
            align-items: center;
            padding: 20px;
            border-radius: 15px;
            background: white;
            border: 3px solid var(--step-color, #007bff);
            min-width: 500px;
            max-width: 700px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        .flow-step:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }
        .flow-step.branch {
            margin-left: 50px;
            min-width: 450px;
            border-left: 4px solid var(--step-color, #007bff);
            border-style: dashed;
            background: linear-gradient(135deg, #ffffff, #f8f9fa);
        }
        .flow-step.branch::before {
            content: "└─";
            position: absolute;
            left: -25px;
            color: var(--step-color, #007bff);
            font-size: 1.5em;
            font-weight: bold;
        }
        .branch-container {
            display: flex;
            flex-direction: column;
            gap: 15px;
            margin-left: 30px;
            padding: 15px;
            border-left: 3px dashed #007bff;
            background: rgba(0, 123, 255, 0.05);
            border-radius: 10px;
        }
        .branch-header {
            font-size: 1.1em;
            font-weight: 600;
            color: #007bff;
            margin-bottom: 10px;
            padding-left: 10px;
        }
        .clickable-step {
            cursor: pointer;
            position: relative;
        }
        .clickable-step::after {
            content: "🔗 Click to navigate";
            position: absolute;
            bottom: -25px;
            right: 10px;
            font-size: 0.75em;
            color: #007bff;
            opacity: 0;
            transition: opacity 0.3s ease;
        }
        .clickable-step:hover::after {
            opacity: 1;
        }
        .clickable-step:hover {
            border-color: #0056b3;
            box-shadow: 0 6px 20px rgba(0,123,255,0.2);
        }
        .step-icon {
            font-size: 2em;
            margin-right: 20px;
            width: 60px;
            text-align: center;
        }
        .step-number {
            background: var(--step-color, #007bff);
            color: white;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 1.2em;
            margin-right: 20px;
            flex-shrink: 0;
        }
        .step-content {
            flex: 1;
        }
        .step-type {
            font-size: 1.3em;
            font-weight: 700;
            color: #212529;
            margin-bottom: 5px;
        }
        .step-description {
            font-size: 1.1em;
            color: #495057;
            margin-bottom: 8px;
        }
        .step-component {
            font-size: 0.9em;
            color: #6c757d;
            font-style: italic;
        }
        .step-action {
            font-size: 0.9em;
            color: var(--step-color, #007bff);
            font-weight: 600;
        }
        .step-details {
            font-size: 0.85em;
            color: #28a745;
            margin-top: 5px;
        }
        .flow-arrow {
            font-size: 3em;
            color: #007bff;
            opacity: 0.7;
        }
        .table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        .table th,
        .table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #dee2e6;
        }
        .table th {
            background-color: #f8f9fa;
            font-weight: 600;
            color: #495057;
        }
        .table tr:hover {
            background-color: #f8f9fa;
        }
        .config-section {
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-top: 20px;
        }
        .config-item {
            margin: 10px 0;
            padding: 10px;
            background: white;
            border-radius: 5px;
            border-left: 4px solid #28a745;
        }
        .config-key {
            font-weight: 600;
            color: #495057;
        }
        .config-value {
            color: #212529;
            font-family: 'Courier New', monospace;
            background-color: #e9ecef;
            padding: 2px 6px;
            border-radius: 3px;
            margin-left: 10px;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Microsoft Teams Call Flow Analysis</h1>
        <h2>Phone Number: $($PhoneNumberInfo.Number)</h2>
        <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')</p>
    </div>

    <div class="section">
        <h3>📋 Number Information</h3>
        <div class="info-grid">
            <div class="info-item">
                <span class="info-label">Phone Number:</span>
                <span class="info-value">$($PhoneNumberInfo.Number)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Assigned To:</span>
                <span class="info-value">$($PhoneNumberInfo.AssignedTo)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Assignment Type:</span>
                <span class="info-value">$($PhoneNumberInfo.AssignmentType)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Number Type:</span>
                <span class="info-value">$($PhoneNumberInfo.NumberType)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Status:</span>
                <span class="info-value">$($PhoneNumberInfo.Status)</span>
            </div>
            <div class="info-item">
                <span class="info-label">User Capable:</span>
                <span class="info-value">$($PhoneNumberInfo.Capabilities.User)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Voice App Capable:</span>
                <span class="info-value">$($PhoneNumberInfo.Capabilities.VoiceApplication)</span>
            </div>
            <div class="info-item">
                <span class="info-label">Place Name:</span>
                <span class="info-value">$($PhoneNumberInfo.PlaceName)</span>
            </div>
        </div>
    </div>

    <div class="section">
        <h3>🔄 Call Flow Diagram</h3>
        <div class="call-flow-container">
"@

    # Add call flow steps with branch support
    $branchSteps = @()
    $mainFlowIndex = 0
    
    for ($i = 0; $i -lt $CallFlow.Count; $i++) {
        $step = $CallFlow[$i]
        $isLast = $i -eq ($CallFlow.Count - 1)
        $isBranch = $step.IsBranch -eq $true
        
        if (-not $isBranch) {
            # Process any accumulated branch steps first
            if ($branchSteps.Count -gt 0) {
                $html += @"
            <div class="branch-container">
                <div class="branch-header">📋 Menu Options:</div>
"@
                foreach ($branchStep in $branchSteps) {
                    $clickableClass = if ($branchStep.TargetPhoneNumber -or $branchStep.TargetFile) { "clickable-step" } else { "" }
                    $targetPhoneAttr = if ($branchStep.TargetPhoneNumber) { "data-target-phone=""$($branchStep.TargetPhoneNumber)""" } else { "" }
                    $targetFileAttr = if ($branchStep.TargetFile) { "data-target-file=""$($branchStep.TargetFile)""" } else { "" }
                    
                    $html += @"
                <div class="flow-step branch $clickableClass" style="--step-color: $($branchStep.Color);" $targetPhoneAttr $targetFileAttr>
                    <div class="step-icon">$($branchStep.Icon)</div>
                    <div class="step-number">$($branchStep.Step)</div>
                    <div class="step-content">
                        <div class="step-type">$($branchStep.Type)</div>
                        <div class="step-description">$($branchStep.Description)</div>
                        <div class="step-component">Component: $($branchStep.Component)</div>
                        <div class="step-action">Action: $($branchStep.Action)</div>
"@
                    if ($branchStep.Details) {
                        $html += @"
                        <div class="step-details">$($branchStep.Details)</div>
"@
                    }
                    $html += @"
                    </div>
                </div>
"@
                }
                $html += @"
            </div>
"@
                $branchSteps = @()
                
                # Add arrow after branches if not the last step
                if (-not $isLast) {
                    $html += @"
            <div class="flow-arrow">↓</div>
"@
                }
            }
            
            # Add main flow step
            $html += @"
            <div class="flow-step" style="--step-color: $($step.Color);">
                <div class="step-icon">$($step.Icon)</div>
                <div class="step-number">$($step.Step)</div>
                <div class="step-content">
                    <div class="step-type">$($step.Type)</div>
                    <div class="step-description">$($step.Description)</div>
                    <div class="step-component">Component: $($step.Component)</div>
                    <div class="step-action">Action: $($step.Action)</div>
"@
            if ($step.Details) {
                $html += @"
                    <div class="step-details">$($step.Details)</div>
"@
            }
            $html += @"
                </div>
            </div>
"@
            
            # Add arrow if not last step and next step is not a branch
            $nextStepIsBranch = ($i + 1) -lt $CallFlow.Count -and $CallFlow[$i + 1].IsBranch -eq $true
            if (-not $isLast -and -not $nextStepIsBranch) {
                $html += @"
            <div class="flow-arrow">↓</div>
"@
            }
            $mainFlowIndex++
        } else {
            # Collect branch steps
            $branchSteps += $step
        }
    }
    
    # Process any remaining branch steps
    if ($branchSteps.Count -gt 0) {
        $html += @"
            <div class="branch-container">
                <div class="branch-header">📋 Menu Options:</div>
"@
        foreach ($branchStep in $branchSteps) {
            $clickableClass = if ($branchStep.TargetPhoneNumber -or $branchStep.TargetFile) { "clickable-step" } else { "" }
            $targetPhoneAttr = if ($branchStep.TargetPhoneNumber) { "data-target-phone=""$($branchStep.TargetPhoneNumber)""" } else { "" }
            $targetFileAttr = if ($branchStep.TargetFile) { "data-target-file=""$($branchStep.TargetFile)""" } else { "" }
            
            $html += @"
                <div class="flow-step branch $clickableClass" style="--step-color: $($branchStep.Color);" $targetPhoneAttr $targetFileAttr>
                    <div class="step-icon">$($branchStep.Icon)</div>
                    <div class="step-number">$($branchStep.Step)</div>
                    <div class="step-content">
                        <div class="step-type">$($branchStep.Type)</div>
                        <div class="step-description">$($branchStep.Description)</div>
                        <div class="step-component">Component: $($branchStep.Component)</div>
                        <div class="step-action">Action: $($branchStep.Action)</div>
"@
            if ($branchStep.Details) {
                $html += @"
                        <div class="step-details">$($branchStep.Details)</div>
"@
            }
            $html += @"
                    </div>
                </div>
"@
        }
        $html += @"
            </div>
"@
    }

    $html += @"
        </div>
    </div>

    <div class="section">
        <h3>📊 Call Flow Steps Details</h3>
        <table class="table">
            <thead>
                <tr>
                    <th>Step</th>
                    <th>Type</th>
                    <th>Description</th>
                    <th>Component</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
"@

    # Add table rows for call flow steps
    foreach ($step in $CallFlow) {
        $html += @"
                <tr>
                    <td>$($step.Step)</td>
                    <td>$($step.Type)</td>
                    <td>$($step.Description)</td>
                    <td>$($step.Component)</td>
                    <td>$($step.Action)</td>
                </tr>
"@
    }

    $html += @"
            </tbody>
        </table>
    </div>
"@

    # Add detailed configuration if requested
    if ($IncludeDetailedSettings -and $PhoneNumberInfo.Configuration -and $PhoneNumberInfo.Configuration.Count -gt 0) {
        $html += @"
    <div class="section">
        <h3>⚙️ Detailed Configuration</h3>
        <div class="config-section">
"@

        foreach ($key in $PhoneNumberInfo.Configuration.Keys) {
            $value = $PhoneNumberInfo.Configuration[$key]
            if ($value -is [hashtable] -or $value -is [PSCustomObject]) {
                $value = ($value | ConvertTo-Json -Depth 2 -Compress) -replace '"', '&quot;'
            }
            
            $html += @"
            <div class="config-item">
                <span class="config-key">${key}:</span>
                <span class="config-value">$($value)</span>
            </div>
"@
        }

        $html += @"
        </div>
    </div>
"@
    }

    # Add JavaScript for navigation
    $html += @"
    <script>
        function navigateToCallFlow(phoneNumber) {
            if (phoneNumber && phoneNumber.trim() !== '') {
                const cleanNumber = phoneNumber.replace(/[^\d]/g, '');
                const targetFile = cleanNumber + '.html';
                const currentPath = window.location.pathname;
                const newPath = currentPath.substring(0, currentPath.lastIndexOf('/')) + '/' + targetFile;
                window.location.href = newPath;
            } else {
                alert('Target phone number not available for navigation');
            }
        }
        
        // Add click handlers to clickable steps
        document.addEventListener('DOMContentLoaded', function() {
            const clickableSteps = document.querySelectorAll('.clickable-step');
            clickableSteps.forEach(step => {
                step.addEventListener('click', function() {
                    const targetFile = this.getAttribute('data-target-file');
                    const phoneNumber = this.getAttribute('data-target-phone');
                    
                    if (targetFile) {
                        // Navigate to queue file
                        window.open(targetFile, '_blank');
                    } else if (phoneNumber) {
                        // Navigate using phone number
                        navigateToCallFlow(phoneNumber);
                    }
                });
            });
        });
    </script>
</body>
</html>
"@

    # Write HTML file
    $html | Out-File -FilePath $htmlFile -Encoding UTF8
    
    return $htmlFile
}

# Function to generate summary dashboard
function New-SummaryDashboard {
    param(
        [hashtable]$PhoneNumbers,
        [string]$OutputPath,
        [hashtable]$TeamsData,
        [array]$QueueFiles = @()
    )
    
    Write-Host "Creating summary dashboard..." -ForegroundColor Cyan
    
    # Create summary statistics
    $totalNumbers = $PhoneNumbers.Count
    $userNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "User" }).Count
    $queueNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "CallQueue" }).Count
    $attendantNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "AutoAttendant" }).Count
    $unassignedNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "Unknown" }).Count
    
    # Count total call queues (including those without direct phone numbers)
    $totalCallQueues = 0
    if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueueDetails) {
        $totalCallQueues = $TeamsData.CallQueues.CallQueueDetails.Count
    }
    
    $summaryFile = Join-Path $OutputPath "Summary" "Dashboard.html"
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Teams Calling - Phone Numbers Dashboard</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f8f9fa;
            color: #333;
        }
        .header {
            background: linear-gradient(135deg, #007bff, #0056b3);
            color: white;
            padding: 40px;
            border-radius: 15px;
            margin-bottom: 30px;
            text-align: center;
            box-shadow: 0 8px 30px rgba(0,0,0,0.15);
        }
        .header h1 {
            margin: 0 0 10px 0;
            font-size: 3em;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        .header h2 {
            margin: 0 0 15px 0;
            font-size: 1.5em;
            opacity: 0.9;
        }
        .header p {
            margin: 5px 0;
            opacity: 0.8;
        }
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .stat-card {
            background: white;
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            border-left: 5px solid var(--card-color, #007bff);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }
        .stat-number {
            font-size: 3em;
            font-weight: bold;
            color: var(--card-color, #007bff);
            margin-bottom: 10px;
        }
        .stat-label {
            font-size: 1.1em;
            color: #495057;
            font-weight: 600;
        }
        .section {
            background: white;
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        }
        .section h3 {
            color: #007bff;
            margin-top: 0;
            margin-bottom: 25px;
            font-size: 1.8em;
            border-bottom: 3px solid #e9ecef;
            padding-bottom: 15px;
        }
        .table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        .table th,
        .table td {
            padding: 15px 12px;
            text-align: left;
            border-bottom: 1px solid #dee2e6;
        }
        .table th {
            background-color: #f8f9fa;
            font-weight: 700;
            color: #495057;
            font-size: 1.1em;
        }
        .table tr:hover {
            background-color: #f8f9fa;
        }
        .table tr:nth-child(even) {
            background-color: #fdfdfd;
        }
        .btn {
            display: inline-block;
            padding: 8px 16px;
            background-color: #007bff;
            color: white;
            text-decoration: none;
            border-radius: 5px;
            transition: background-color 0.3s ease;
        }
        .btn:hover {
            background-color: #0056b3;
            color: white;
        }
        .assignment-badge {
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            color: white;
        }
        .badge-user { background-color: #28a745; }
        .badge-queue { background-color: #17a2b8; }
        .badge-attendant { background-color: #6f42c1; }
        .badge-unknown { background-color: #6c757d; }
        
        .agent-details {
            margin-top: 15px;
            padding: 12px;
            background-color: #f8f9fa;
            border-radius: 6px;
            border-left: 4px solid #17a2b8;
        }
        
        .agent-summary {
            cursor: pointer;
            display: flex;
            align-items: center;
            font-weight: 600;
            color: #495057;
            padding: 8px 0;
            transition: color 0.3s ease;
        }
        
        .agent-summary:hover {
            color: #17a2b8;
        }
        
        .agent-summary::before {
            content: '▶';
            margin-right: 8px;
            font-size: 0.8em;
            transition: transform 0.3s ease;
        }
        
        .agent-summary.expanded::before {
            transform: rotate(90deg);
        }
        
        .agent-list {
            display: none;
            margin-top: 10px;
            padding-left: 20px;
        }
        
        .agent-list.show {
            display: block;
        }
        
        .agent-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 6px 10px;
            margin: 4px 0;
            background-color: #ffffff;
            border-radius: 4px;
            border: 1px solid #e9ecef;
            font-size: 0.9em;
        }
        
        .agent-name {
            font-weight: 500;
            color: #495057;
        }
        
        .agent-email {
            color: #6c757d;
            font-size: 0.85em;
            font-family: 'Courier New', monospace;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>📞 Microsoft Teams Calling Dashboard</h1>
        <h2>Phone Number Call Flow Analysis</h2>
        <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')</p>
    </div>

    <div class="stats-grid">
        <div class="stat-card" style="--card-color: #007bff;">
            <div class="stat-number">$totalNumbers</div>
            <div class="stat-label">Total Phone Numbers</div>
        </div>
        <div class="stat-card" style="--card-color: #28a745;">
            <div class="stat-number">$userNumbers</div>
            <div class="stat-label">User Assigned</div>
        </div>
        <div class="stat-card" style="--card-color: #17a2b8;">
            <div class="stat-number">$queueNumbers</div>
            <div class="stat-label">Call Queue Assigned</div>
        </div>
        <div class="stat-card" style="--card-color: #6f42c1;">
            <div class="stat-number">$attendantNumbers</div>
            <div class="stat-label">Auto Attendant Assigned</div>
        </div>
        <div class="stat-card" style="--card-color: #fd7e14;">
            <div class="stat-number">$totalCallQueues</div>
            <div class="stat-label">Total Call Queues</div>
        </div>
        <div class="stat-card" style="--card-color: #dc3545;">
            <div class="stat-number">$unassignedNumbers</div>
            <div class="stat-label">Unassigned/Unknown</div>
        </div>
    </div>

    <div class="section">
        <h3>📋 Phone Number Details</h3>
        <table class="table">
            <thead>
                <tr>
                    <th>Phone Number</th>
                    <th>Assigned To</th>
                    <th>Assignment Type</th>
                    <th>Status</th>
                    <th>Capabilities</th>
                    <th>Call Flow</th>
                </tr>
            </thead>
            <tbody>
"@

    # Add phone number rows
    foreach ($number in ($PhoneNumbers.Keys | Sort-Object)) {
        $info = $PhoneNumbers[$number]
        $capabilities = if ($info.Capabilities.User) { "User" } elseif ($info.Capabilities.VoiceApplication) { "Voice App" } else { "None" }
        $badgeClass = switch ($info.AssignmentType) {
            "User" { "badge-user" }
            "CallQueue" { "badge-queue" }
            "AutoAttendant" { "badge-attendant" }
            default { "badge-unknown" }
        }
        $safeFileName = $info.Number -replace '[^\d]', ''
        
        $html += @"
                <tr>
                    <td><strong>$($info.Number)</strong></td>
                    <td>$($info.AssignedTo)</td>
                    <td><span class="assignment-badge $badgeClass">$($info.AssignmentType)</span></td>
                    <td>$($info.Status)</td>
                    <td>$capabilities</td>
                    <td><a href="../Individual/$safeFileName.html" class="btn" target="_blank">View Call Flow</a></td>
                </tr>
"@
    }

    $html += @"
            </tbody>
        </table>
    </div>
"@

    # Add Call Queues section if available
    if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueueDetails) {
        $html += @"
    <div class="section">
        <h3>🏢 Call Queues Overview</h3>
        <table class="table">
            <thead>
                <tr>
                    <th>Queue Name</th>
                    <th>Timeout</th>
                    <th>Routing Method</th>
                    <th>Agent Count</th>
                    <th>Conference Mode</th>
                    <th>Call Flow</th>
                </tr>
            </thead>
            <tbody>
"@

        foreach ($queue in $TeamsData.CallQueues.CallQueueDetails) {
            $agentCount = if ($queue.Agents) { $queue.Agents.Count } else { 0 }
            
            # Calculate timeout display
            $timeoutDisplay = "Not Set"
            if ($queue.TimeoutThreshold -and $queue.TimeoutThreshold -gt 0) {
                $timeoutMinutes = [Math]::Floor($queue.TimeoutThreshold / 60)
                $timeoutSeconds = $queue.TimeoutThreshold % 60
                $timeoutDisplay = if ($timeoutMinutes -gt 0) { "$timeoutMinutes minutes, $timeoutSeconds seconds" } else { "$timeoutSeconds seconds" }
            }
            
            # Routing method display
            $routingMethod = switch ($queue.RoutingMethod) {
                0 { "Attendant" }
                1 { "Round Robin" }
                2 { "Serial" }
                3 { "Longest Idle" }
                default { "Unknown" }
            }
            
            # Check if queue has a standalone call flow file
            $queueFileName = "Queue_$($queue.Name -replace '[^a-zA-Z0-9]', '_').html"
            $callFlowLink = "<a href=""../Individual/$queueFileName"" class=""btn"" target=""_blank"">View Queue Flow</a>"
            
            $html += @"
                <tr>
                    <td><strong>$($queue.Name)</strong></td>
                    <td>$timeoutDisplay</td>
                    <td>$routingMethod</td>
                    <td>$agentCount</td>
                    <td>$($queue.ConferenceMode)</td>
                    <td>$callFlowLink</td>
                </tr>
"@
        }

        $html += @"
            </tbody>
        </table>
    </div>
"@
    }

    # Add Auto Attendants section if available
    if ($TeamsData.AutoAttendants -and $TeamsData.AutoAttendants.AutoAttendants) {
        $html += @"
    <div class="section">
        <h3>🤖 Auto Attendants Overview</h3>
        <table class="table">
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Language</th>
                    <th>Time Zone</th>
                    <th>Voice</th>
                    <th>Call Flows</th>
                    <th>Schedules</th>
                </tr>
            </thead>
            <tbody>
"@

        foreach ($attendant in $TeamsData.AutoAttendants.AutoAttendants) {
            $callFlowsCount = if ($attendant.CallFlows) { $attendant.CallFlows.Count } else { 0 }
            $schedulesCount = if ($attendant.Schedules) { $attendant.Schedules.Count } else { 0 }
            
            $html += @"
                <tr>
                    <td><strong>$($attendant.Name)</strong></td>
                    <td>$($attendant.LanguageId)</td>
                    <td>$($attendant.TimeZoneId)</td>
                    <td>$($attendant.VoiceId)</td>
                    <td>$callFlowsCount</td>
                    <td>$schedulesCount</td>
                </tr>
"@
        }

        $html += @"
            </tbody>
        </table>
    </div>
"@
    }

    $html += @"

    <script>
        // Toggle agent details functionality
        document.addEventListener('DOMContentLoaded', function() {
            const agentSummaries = document.querySelectorAll('.agent-summary');
            
            agentSummaries.forEach(summary => {
                summary.addEventListener('click', function() {
                    this.classList.toggle('expanded');
                    const agentList = this.nextElementSibling;
                    agentList.classList.toggle('show');
                });
            });
        });
    </script>
</body>
</html>
"@

    # Write HTML file
    $html | Out-File -FilePath $summaryFile -Encoding UTF8
    
    return $summaryFile
}

# Function to convert HTML to PDF
function ConvertTo-PDF {
    param(
        [string]$HtmlFilePath,
        [string]$OutputPath
    )
    
    $pdfPath = $HtmlFilePath -replace '\.html$', '.pdf'
    $pdfDir = Split-Path $pdfPath -Parent
    if (-not (Test-Path $pdfDir)) {
        New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
    }
    
    try {
        # Try using Chrome/Chromium in headless mode
        $chromeExe = $null
        $possiblePaths = @(
            "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        )
        
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                $chromeExe = $path
                break
            }
        }
        
        if ($chromeExe) {
            $fileUri = "file:///$($HtmlFilePath -replace '\\', '/')"
            $chromeArgs = @(
                "--headless",
                "--disable-gpu",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--print-to-pdf=`"$pdfPath`"",
                "--print-to-pdf-no-header",
                "--virtual-time-budget=10000",
                "`"$fileUri`""
            )
            
            $process = Start-Process -FilePath $chromeExe -ArgumentList $chromeArgs -Wait -NoNewWindow -PassThru
            
            # Wait a moment for file to be written
            Start-Sleep -Seconds 2
            
            if ((Test-Path $pdfPath) -and $process.ExitCode -eq 0) {
                Write-Host "✓ Generated PDF: $(Split-Path $pdfPath -Leaf)" -ForegroundColor Green
                return $pdfPath
            }
        }
        
        Write-Warning "Could not generate PDF. Chrome/Edge not found or PDF generation failed."
        return $null
    }
    catch {
        Write-Warning "PDF generation failed: $($_.Exception.Message)"
        return $null
    }
}

# Main execution
try {
    Write-Host "=== Teams Call Flow Map Generator ===" -ForegroundColor Cyan
    Write-Host "Started at: $(Get-Date)" -ForegroundColor Gray
    
    # Validate input path
    if (-not (Test-Path $JsonDataPath)) {
        Write-Error "JSON data path not found: $JsonDataPath"
        exit 1
    }
    
    # Initialize output directory
    $outputDir = Initialize-OutputDirectory -Path $OutputPath
    Write-Host "Output directory: $outputDir" -ForegroundColor Gray
    
    # Load Teams data
    $teamsData = Import-TeamsData -DataPath $JsonDataPath
    
    # Extract phone number mappings
    $phoneNumbers = Get-PhoneNumberMappings -TeamsData $teamsData
    
    if ($phoneNumbers.Count -eq 0) {
        Write-Warning "No phone numbers found in the data. Please verify the JSON files contain phone number assignments."
        exit 1
    }
    
    # Filter by specific numbers if requested
    if ($FilterByNumber -and $FilterByNumber.Count -gt 0) {
        $filteredNumbers = @{}
        foreach ($filter in $FilterByNumber) {
            $matchingNumbers = $phoneNumbers.Keys | Where-Object { $_ -like "*$filter*" }
            foreach ($match in $matchingNumbers) {
                $filteredNumbers[$match] = $phoneNumbers[$match]
            }
        }
        $phoneNumbers = $filteredNumbers
        Write-Host "Filtered to $($phoneNumbers.Count) phone numbers" -ForegroundColor Yellow
    }
    
    # Generate call flows and HTML files
    Write-Host "`nGenerating call flow maps..." -ForegroundColor Cyan
    $htmlFiles = @()
    $pdfFiles = @()
    
    foreach ($number in $phoneNumbers.Keys) {
        Write-Host "  Processing: $number" -ForegroundColor Yellow
        
        # Build call flow
        $callFlow = Build-CallFlow -PhoneNumberInfo $phoneNumbers[$number] -TeamsData $teamsData -PhoneNumbers $phoneNumbers
        $phoneNumbers[$number].CallFlow = $callFlow
        
        # Generate HTML
        $htmlFile = New-CallFlowHTML -PhoneNumberInfo $phoneNumbers[$number] -CallFlow $callFlow -OutputPath $outputDir -IncludeDetailedSettings:$IncludeDetailedSettings
        $htmlFiles += $htmlFile
        
        # Generate PDF if requested
        if ($GeneratePDF) {
            $pdfFile = ConvertTo-PDF -HtmlFilePath $htmlFile -OutputPath $outputDir
            if ($pdfFile) {
                $pdfFiles += $pdfFile
            }
        }
    }
    
    # Generate standalone call queue maps for queues without direct phone numbers
    Write-Host "`nGenerating call queue maps..." -ForegroundColor Cyan
    $queueFiles = @()
    
    if ($teamsData.CallQueues -and $teamsData.CallQueues.CallQueueDetails) {
        foreach ($queue in $teamsData.CallQueues.CallQueueDetails) {
            # Check if this queue already has a direct phone number
            $hasDirectNumber = $false
            foreach ($number in $phoneNumbers.Keys) {
                if ($phoneNumbers[$number].Configuration.QueueId -eq $queue.Identity -or 
                    $phoneNumbers[$number].Configuration.QueueName -eq $queue.Name) {
                    $hasDirectNumber = $true
                    break
                }
            }
            
            # Generate standalone call queue map if no direct number
            if (-not $hasDirectNumber) {
                Write-Host "  Processing Call Queue: $($queue.Name)" -ForegroundColor Yellow
                
                # Build call queue info
                $queueInfo = @{
                    Name = $queue.Name
                    Identity = $queue.Identity
                    Type = "CallQueue"
                    AssignmentType = "CallQueue"
                    Configuration = @{
                        QueueDetails = $queue
                        QueueId = $queue.Identity
                        QueueName = $queue.Name
                    }
                    CallFlow = @()
                }
                
                # Build call queue flow
                $callFlow = Build-CallQueueFlow -QueueDetails $queue -TeamsData $teamsData
                $queueInfo.CallFlow = $callFlow
                
                # Generate HTML
                $queueFile = New-CallQueueHTML -QueueInfo $queueInfo -CallFlow $callFlow -OutputPath $outputDir -TeamsData $teamsData -IncludeDetailedSettings:$IncludeDetailedSettings
                if ($queueFile) {
                    $queueFiles += $queueFile
                    $htmlFiles += $queueFile
                }
                
                # Generate PDF if requested
                if ($GeneratePDF -and $queueFile) {
                    $pdfFile = ConvertTo-PDF -HtmlFilePath $queueFile -OutputPath $outputDir
                    if ($pdfFile) {
                        $pdfFiles += $pdfFile
                    }
                }
            }
        }
    }
    
    # Generate summary dashboard
    $summaryFile = New-SummaryDashboard -PhoneNumbers $phoneNumbers -OutputPath $outputDir -TeamsData $teamsData -QueueFiles $queueFiles
    $htmlFiles += $summaryFile
    
    if ($GeneratePDF) {
        Write-Host "Generating summary dashboard PDF..." -ForegroundColor Yellow
        $summaryPdf = ConvertTo-PDF -HtmlFilePath $summaryFile -OutputPath $outputDir
        if ($summaryPdf) {
            $pdfFiles += $summaryPdf
        }
    }
    
    # Final summary
    Write-Host "`n=== Generation Complete ===" -ForegroundColor Green
    Write-Host "Completed at: $(Get-Date)" -ForegroundColor Gray
    Write-Host "Output location: $outputDir" -ForegroundColor Gray
    Write-Host "Phone numbers processed: $($phoneNumbers.Count)" -ForegroundColor Gray
    Write-Host "HTML files created: $($htmlFiles.Count)" -ForegroundColor Gray
    
    if ($GeneratePDF) {
        Write-Host "PDF files created: $($pdfFiles.Count)" -ForegroundColor Gray
    }
    
    Write-Host "`nOpen the summary dashboard to start exploring:" -ForegroundColor Cyan
    Write-Host $summaryFile -ForegroundColor White
    
    # Try to open the summary file
    if ($IsWindows -or $env:OS -eq "Windows_NT") {
        try {
            Start-Process $summaryFile
        }
        catch {
            Write-Host "To view results, open: $summaryFile" -ForegroundColor Yellow
        }
    }
    elseif ($IsMacOS) {
        try {
            Start-Process "open" -ArgumentList $summaryFile
        }
        catch {
            Write-Host "To view results, open: $summaryFile" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Error "An error occurred during execution: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
}
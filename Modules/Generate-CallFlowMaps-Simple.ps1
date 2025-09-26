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
    Author: Generated for TCP Calling Issues Analysis
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
        [hashtable]$TeamsData
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
                    }
                    
                    if ($attendantDetail.CallHandlingAssociations -and $attendantDetail.CallHandlingAssociations.Count -gt 0) {
                        $callFlow += @{
                            Step = 5
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

    # Add call flow steps
    for ($i = 0; $i -lt $CallFlow.Count; $i++) {
        $step = $CallFlow[$i]
        $isLast = $i -eq ($CallFlow.Count - 1)
        
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

    $html += @"
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
        [hashtable]$TeamsData
    )
    
    Write-Host "Creating summary dashboard..." -ForegroundColor Cyan
    
    # Create summary statistics
    $totalNumbers = $PhoneNumbers.Count
    $userNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "User" }).Count
    $queueNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "CallQueue" }).Count
    $attendantNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "AutoAttendant" }).Count
    $unassignedNumbers = ($PhoneNumbers.Values | Where-Object { $_.AssignmentType -eq "Unknown" }).Count
    
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
    if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueues) {
        $html += @"
    <div class="section">
        <h3>🏢 Call Queues Overview</h3>
        <table class="table">
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Routing Method</th>
                    <th>Conference Mode</th>
                    <th>Agent Count</th>
                    <th>Has Overflow</th>
                    <th>Has Timeout</th>
                </tr>
            </thead>
            <tbody>
"@

        foreach ($queue in $TeamsData.CallQueues.CallQueues) {
            $agentCount = if ($queue.Users) { $queue.Users.Count } else { 0 }
            $hasOverflow = if ($queue.OverflowAction) { "✅" } else { "❌" }
            $hasTimeout = if ($queue.TimeoutAction) { "✅" } else { "❌" }
            
            $html += @"
                <tr>
                    <td><strong>$($queue.Name)</strong></td>
                    <td>$($queue.RoutingMethod)</td>
                    <td>$($queue.ConferenceMode)</td>
                    <td>$agentCount</td>
                    <td>$hasOverflow</td>
                    <td>$hasTimeout</td>
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
        $callFlow = Build-CallFlow -PhoneNumberInfo $phoneNumbers[$number] -TeamsData $teamsData
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
    
    # Generate summary dashboard
    $summaryFile = New-SummaryDashboard -PhoneNumbers $phoneNumbers -OutputPath $outputDir -TeamsData $teamsData
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
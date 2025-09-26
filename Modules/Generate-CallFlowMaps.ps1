#Requires -Modules @{ModuleName="PSWriteHTML"; ModuleVersion="0.0.1"}
<#
.SYNOPSIS
    Generates call flow maps and diagrams for each phone number in the Teams tenant
    
.DESCRIPTION
    This script processes the JSON files created by Gather-Teams-Calling-Data.ps1 and creates:
    - Individual call flow maps for each phone number
    - Visual diagrams showing call routing and settings
    - PDF reports with comprehensive call flow analysis
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
    .\Generate-CallFlowMaps.ps1 -JsonDataPath ".\TeamsCallingData_20250926_110848" -GeneratePDF
    
.EXAMPLE
    .\Generate-CallFlowMaps.ps1 -JsonDataPath "C:\TeamsData" -OutputPath "C:\CallFlowMaps" -IncludeDetailedSettings -GeneratePDF
    
.NOTES
    Author: Generated for TCP Calling Issues Analysis
    Date: September 26, 2025
    Requires: PSWriteHTML module for HTML/PDF generation
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

# Check for required modules
function Test-RequiredModules {
    $requiredModules = @("PSWriteHTML")
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Warning "Missing required modules: $($missingModules -join ', ')"
        Write-Host "Installing missing modules..." -ForegroundColor Yellow
        
        foreach ($module in $missingModules) {
            try {
                Install-Module -Name $module -Force -Scope CurrentUser
                Write-Host "✓ Installed: $module" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install $module`: $($_.Exception.Message)"
                return $false
            }
        }
    }
    
    # Import modules
    foreach ($module in $requiredModules) {
        try {
            Import-Module -Name $module -Force
        }
        catch {
            Write-Error "Failed to import $module`: $($_.Exception.Message)"
            return $false
        }
    }
    
    return $true
}

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
    $subDirs = @("HTML", "PDF", "Resources", "Individual", "Summary")
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
            }
            
            $callFlow += @{
                Step = 2
                Type = "User Routing"
                Description = "Call routed to user: $($PhoneNumberInfo.AssignedTo)"
                Details = "Enterprise Voice: $($PhoneNumberInfo.Configuration.EnterpriseVoiceEnabled)"
                Component = "Teams Client"
                Action = "Ring user devices"
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
            }
            
            $callFlow += @{
                Step = 2
                Type = "Call Queue Processing"
                Description = "Call Queue: $($PhoneNumberInfo.Configuration.QueueName)"
                Details = "Queue ID: $($PhoneNumberInfo.Configuration.QueueId)"
                Component = "Call Queue Engine"
                Action = "Queue management"
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
                    }
                    
                    if ($queueDetail.OverflowAction) {
                        $callFlow += @{
                            Step = 4
                            Type = "Overflow Action"
                            Description = "Action: $($queueDetail.OverflowAction)"
                            Details = "Threshold: $($queueDetail.OverflowThreshold) calls"
                            Component = "Queue Logic"
                            Action = "Handle overflow"
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
            }
            
            $callFlow += @{
                Step = 2
                Type = "Auto Attendant Processing"
                Description = "Auto Attendant: $($PhoneNumberInfo.Configuration.AttendantName)"
                Details = "Attendant ID: $($PhoneNumberInfo.Configuration.AttendantId)"
                Component = "Auto Attendant Engine"
                Action = "Process call flow"
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
                    }
                    
                    if ($attendantDetail.DefaultCallFlow -and $attendantDetail.DefaultCallFlow.Menu) {
                        $menuOptions = $attendantDetail.DefaultCallFlow.Menu.MenuOptions.Count
                        $callFlow += @{
                            Step = 4
                            Type = "Menu Presentation"
                            Description = "Present menu with $menuOptions options"
                            Details = "Default call flow: $($attendantDetail.DefaultCallFlow.Name)"
                            Component = "Menu System"
                            Action = "Play menu options"
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
    
    $htmlContent = New-HTML -TitleText "Call Flow: $($PhoneNumberInfo.Number)" -Online {
        New-HTMLHeader {
            New-HTMLText -Text "Microsoft Teams Call Flow Analysis" -FontSize 24 -FontWeight bold -Color Blue
            New-HTMLText -Text "Phone Number: $($PhoneNumberInfo.Number)" -FontSize 18 -FontWeight bold
            New-HTMLText -Text "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -FontSize 12 -Color Gray
        }
        
        New-HTMLSection -HeaderText "Number Information" -CanCollapse {
            New-HTMLPanel {
                New-HTMLTable -DataTable @(
                    [PSCustomObject]@{
                        Property = "Phone Number"
                        Value = $PhoneNumberInfo.Number
                    },
                    [PSCustomObject]@{
                        Property = "Assigned To"
                        Value = $PhoneNumberInfo.AssignedTo
                    },
                    [PSCustomObject]@{
                        Property = "Assignment Type"
                        Value = $PhoneNumberInfo.AssignmentType
                    },
                    [PSCustomObject]@{
                        Property = "Number Type"
                        Value = $PhoneNumberInfo.NumberType
                    },
                    [PSCustomObject]@{
                        Property = "Status"
                        Value = $PhoneNumberInfo.Status
                    },
                    [PSCustomObject]@{
                        Property = "User Capable"
                        Value = $PhoneNumberInfo.Capabilities.User
                    },
                    [PSCustomObject]@{
                        Property = "Voice App Capable"
                        Value = $PhoneNumberInfo.Capabilities.VoiceApplication
                    }
                ) -HideFooter
            }
        }
        
        New-HTMLSection -HeaderText "Call Flow Diagram" {
            New-HTMLPanel {
                # Create flowchart using HTML/CSS
                New-HTMLTag -Tag "div" -Attributes @{ class = "call-flow-container" } {
                    for ($i = 0; $i -lt $CallFlow.Count; $i++) {
                        $step = $CallFlow[$i]
                        $isLast = $i -eq ($CallFlow.Count - 1)
                        
                        New-HTMLTag -Tag "div" -Attributes @{ class = "flow-step" } {
                            New-HTMLTag -Tag "div" -Attributes @{ class = "step-number" } {
                                New-HTMLText -Text $step.Step
                            }
                            New-HTMLTag -Tag "div" -Attributes @{ class = "step-content" } {
                                New-HTMLText -Text $step.Type -FontWeight bold -FontSize 14
                                New-HTMLText -Text $step.Description -FontSize 12
                                New-HTMLText -Text "Component: $($step.Component)" -FontSize 10 -Color Gray
                                New-HTMLText -Text "Action: $($step.Action)" -FontSize 10 -Color DarkBlue
                                if ($step.Details) {
                                    New-HTMLText -Text $step.Details -FontSize 10 -Color DarkGreen
                                }
                            }
                        }
                        
                        if (-not $isLast) {
                            New-HTMLTag -Tag "div" -Attributes @{ class = "flow-arrow" } {
                                New-HTMLText -Text "↓"
                            }
                        }
                    }
                }
            }
        }
        
        New-HTMLSection -HeaderText "Call Flow Steps" -CanCollapse {
            New-HTMLTable -DataTable $CallFlow -HideFooter
        }
        
        if ($IncludeDetailedSettings -and $PhoneNumberInfo.Configuration) {
            New-HTMLSection -HeaderText "Detailed Configuration" -CanCollapse {
                $configTable = @()
                foreach ($key in $PhoneNumberInfo.Configuration.Keys) {
                    $value = $PhoneNumberInfo.Configuration[$key]
                    if ($value -is [hashtable] -or $value -is [PSCustomObject]) {
                        $value = $value | ConvertTo-Json -Depth 3 -Compress
                    }
                    $configTable += [PSCustomObject]@{
                        Setting = $key
                        Value = $value
                    }
                }
                New-HTMLTable -DataTable $configTable -HideFooter
            }
        }
        
        # Add custom CSS for call flow diagram
        New-HTMLTag -Tag "style" {
            @"
            .call-flow-container {
                display: flex;
                flex-direction: column;
                align-items: center;
                padding: 20px;
            }
            .flow-step {
                display: flex;
                align-items: center;
                margin: 10px 0;
                padding: 15px;
                border: 2px solid #0078d4;
                border-radius: 10px;
                background: linear-gradient(145deg, #f0f8ff, #e6f3ff);
                min-width: 400px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            }
            .step-number {
                background: #0078d4;
                color: white;
                border-radius: 50%;
                width: 30px;
                height: 30px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-weight: bold;
                margin-right: 15px;
                flex-shrink: 0;
            }
            .step-content {
                flex: 1;
            }
            .flow-arrow {
                font-size: 24px;
                color: #0078d4;
                margin: 5px 0;
            }
"@
        }
    }
    
    $htmlFile = Join-Path $OutputPath "Individual" "$($PhoneNumberInfo.Number -replace '[^\d]', '').html"
    $htmlContent | Out-File -FilePath $htmlFile -Encoding UTF8
    
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
    
    $htmlContent = New-HTML -TitleText "Teams Calling - Phone Numbers Dashboard" -Online {
        New-HTMLHeader {
            New-HTMLText -Text "Microsoft Teams Calling Dashboard" -FontSize 28 -FontWeight bold -Color Blue
            New-HTMLText -Text "Phone Number Call Flow Analysis" -FontSize 18
            New-HTMLText -Text "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -FontSize 12 -Color Gray
        }
        
        New-HTMLSection -HeaderText "Summary Statistics" {
            New-HTMLPanel {
                New-HTMLChart -Type Bar -Title "Phone Number Distribution" {
                    New-HTMLChartBarOptions -Type barStacked
                    New-HTMLChartLegend -Names @('User Numbers', 'Call Queue Numbers', 'Auto Attendant Numbers', 'Unassigned Numbers')
                    New-HTMLChartBar -Name 'Phone Numbers' -Value @($userNumbers, $queueNumbers, $attendantNumbers, $unassignedNumbers)
                }
            }
            
            New-HTMLPanel {
                New-HTMLTable -DataTable @(
                    [PSCustomObject]@{ Metric = "Total Phone Numbers"; Count = $totalNumbers },
                    [PSCustomObject]@{ Metric = "User Assigned"; Count = $userNumbers },
                    [PSCustomObject]@{ Metric = "Call Queue Assigned"; Count = $queueNumbers },
                    [PSCustomObject]@{ Metric = "Auto Attendant Assigned"; Count = $attendantNumbers },
                    [PSCustomObject]@{ Metric = "Unassigned/Unknown"; Count = $unassignedNumbers }
                ) -HideFooter
            }
        }
        
        New-HTMLSection -HeaderText "Phone Number Details" {
            $phoneNumberTable = @()
            foreach ($number in $PhoneNumbers.Keys | Sort-Object) {
                $info = $PhoneNumbers[$number]
                $phoneNumberTable += [PSCustomObject]@{
                    PhoneNumber = $info.Number
                    AssignedTo = $info.AssignedTo
                    AssignmentType = $info.AssignmentType
                    Status = $info.Status
                    Capabilities = if ($info.Capabilities.User) { "User" } elseif ($info.Capabilities.VoiceApplication) { "Voice App" } else { "None" }
                    CallFlowLink = "<a href='Individual/$($info.Number -replace '[^\d]', '').html' target='_blank'>View Call Flow</a>"
                }
            }
            New-HTMLTable -DataTable $phoneNumberTable -HideFooter
        }
        
        if ($TeamsData.CallQueues -and $TeamsData.CallQueues.CallQueues) {
            New-HTMLSection -HeaderText "Call Queues Overview" -CanCollapse {
                $queueTable = @()
                foreach ($queue in $TeamsData.CallQueues.CallQueues) {
                    $queueTable += [PSCustomObject]@{
                        Name = $queue.Name
                        Identity = $queue.Identity
                        RoutingMethod = $queue.RoutingMethod
                        ConferenceMode = $queue.ConferenceMode
                        AgentCount = if ($queue.Users) { $queue.Users.Count } else { 0 }
                        HasOverflowAction = [bool]$queue.OverflowAction
                        HasTimeoutAction = [bool]$queue.TimeoutAction
                    }
                }
                New-HTMLTable -DataTable $queueTable -HideFooter
            }
        }
        
        if ($TeamsData.AutoAttendants -and $TeamsData.AutoAttendants.AutoAttendants) {
            New-HTMLSection -HeaderText "Auto Attendants Overview" -CanCollapse {
                $attendantTable = @()
                foreach ($attendant in $TeamsData.AutoAttendants.AutoAttendants) {
                    $attendantTable += [PSCustomObject]@{
                        Name = $attendant.Name
                        Identity = $attendant.Identity
                        LanguageId = $attendant.LanguageId
                        TimeZoneId = $attendant.TimeZoneId
                        VoiceId = $attendant.VoiceId
                        CallFlowsCount = if ($attendant.CallFlows) { $attendant.CallFlows.Count } else { 0 }
                        SchedulesCount = if ($attendant.Schedules) { $attendant.Schedules.Count } else { 0 }
                    }
                }
                New-HTMLTable -DataTable $attendantTable -HideFooter
            }
        }
    }
    
    $summaryFile = Join-Path $OutputPath "Summary" "Dashboard.html"
    $htmlContent | Out-File -FilePath $summaryFile -Encoding UTF8
    
    return $summaryFile
}

# Function to convert HTML to PDF
function ConvertTo-PDF {
    param(
        [string]$HtmlFilePath,
        [string]$OutputPath
    )
    
    $pdfPath = $HtmlFilePath -replace '\.html$', '.pdf'
    
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
            $chromeArgs = @(
                "--headless",
                "--disable-gpu",
                "--no-sandbox",
                "--print-to-pdf=`"$pdfPath`"",
                "--print-to-pdf-no-header",
                "`"file:///$($HtmlFilePath -replace '\\', '/')`""
            )
            
            Start-Process -FilePath $chromeExe -ArgumentList $chromeArgs -Wait -NoNewWindow
            
            if (Test-Path $pdfPath) {
                Write-Host "✓ Generated PDF: $pdfPath" -ForegroundColor Green
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
    
    # Check for required modules
    if (-not (Test-RequiredModules)) {
        Write-Error "Failed to install/import required modules. Exiting."
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
}
catch {
    Write-Error "An error occurred during execution: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
}
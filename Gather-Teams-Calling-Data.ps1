#Requires -Modules MicrosoftTeams
<#
.SYNOPSIS
    Master script to gather all Microsoft Teams calling data and export to JSON
    
.DESCRIPTION
    This script orchestrates the collection of all Teams calling-related settings including:
    - Calling policies and configurations
    - Call queues and auto attendants
    - User calling settings and assignments
    - Emergency calling locations and policies
    - Voice routing and PSTN settings
    - Compliance and security settings
    
.PARAMETER OutputPath
    Path where JSON files will be saved. Defaults to current directory with timestamp folder
    
.PARAMETER TenantId
    Optional tenant ID for authentication
    
.PARAMETER IncludeUserData
    Switch to include detailed user calling data (may take longer for large tenants)
    
.PARAMETER CompressOutput
    Switch to compress the output into a ZIP file
    
.PARAMETER GenerateCallFlowMaps
    Switch to automatically generate call flow maps after data collection
    
.PARAMETER GeneratePDF
    Switch to generate PDF files (requires Python with Playwright)
    
.PARAMETER CallFlowFilterNumbers
    Optional filter for call flow generation - only generate maps for specific numbers
    
.EXAMPLE
    .\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\TeamsData" -IncludeUserData -GenerateCallFlowMaps -GeneratePDF
    
.EXAMPLE
    .\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps -CallFlowFilterNumbers "+1984308", "+1800" -GeneratePDF
    
.NOTES
    Author: Generated for TCP Calling Issues Analysis
    Date: September 26, 2025
    Requires: MicrosoftTeams PowerShell module and appropriate admin permissions
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeUserData,
    
    [Parameter(Mandatory = $false)]
    [switch]$CompressOutput,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateCallFlowMaps,
    
    [Parameter(Mandatory = $false)]
    [switch]$GeneratePDF,
    
    [Parameter(Mandatory = $false)]
    [string[]]$CallFlowFilterNumbers
)

# Set error action preference
$ErrorActionPreference = "Continue"
$WarningPreference = "Continue"

# Import required modules
$ModulesPath = Join-Path $PSScriptRoot "Modules"

# Define the modules to load
$ModuleList = @(
    "TeamsCallingPolicies.psm1",
    "TeamsCallQueues.psm1",
    "TeamsAutoAttendants.psm1",
    "TeamsUserSettings.psm1",
    "TeamsEmergencyLocations.psm1",
    "TeamsVoiceRouting.psm1",
    "TeamsComplianceSettings.psm1"
)

# Function to initialize output directory
function Initialize-OutputDirectory {
    param([string]$Path)
    
    if (-not $Path) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $Path = Join-Path $PWD "TeamsCallingData_$timestamp"
    }
    
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    
    return $Path
}

# Function to test Teams connection
function Test-TeamsConnection {
    try {
        $context = Get-CsTeamsEnvironment -ErrorAction SilentlyContinue
        if ($context) {
            Write-Host "✓ Connected to Teams PowerShell" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "✗ Not connected to Teams PowerShell" -ForegroundColor Red
        return $false
    }
    return $false
}

# Function to connect to Teams
function Connect-TeamsService {
    param([string]$TenantId)
    
    Write-Host "Connecting to Microsoft Teams PowerShell..." -ForegroundColor Cyan
    
    try {
        if ($TenantId) {
            Connect-MicrosoftTeams -TenantId $TenantId
        } else {
            Connect-MicrosoftTeams
        }
        
        if (Test-TeamsConnection) {
            Write-Host "Successfully connected to Microsoft Teams" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "Failed to connect to Microsoft Teams: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    return $false
}

# Function to load custom modules
function Import-CustomModules {
    param([string[]]$Modules, [string]$ModulesPath)
    
    $loadedModules = @()
    
    foreach ($module in $Modules) {
        $modulePath = Join-Path $ModulesPath $module
        
        if (Test-Path $modulePath) {
            try {
                Import-Module $modulePath -Force
                $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($module)
                $loadedModules += $moduleName
                Write-Host "✓ Loaded module: $moduleName" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to load module $module`: $($_.Exception.Message)"
            }
        }
        else {
            Write-Warning "Module not found: $modulePath"
        }
    }
    
    return $loadedModules
}

# Function to export data to JSON
function Export-DataToJson {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Data,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )
    
    try {
        $jsonPath = Join-Path $OutputPath "$FileName.json"
        $Data | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-Host "✓ Exported: $FileName.json" -ForegroundColor Green
        return $jsonPath
    }
    catch {
        Write-Error "Failed to export $FileName`: $($_.Exception.Message)"
        return $null
    }
}

# Function to create summary report
function New-SummaryReport {
    param(
        [hashtable]$CollectedData,
        [string]$OutputPath
    )
    
    $summary = @{
        CollectionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
        TenantInfo = @{
            TenantId = (Get-CsTenant -ErrorAction SilentlyContinue).TenantId
            DisplayName = (Get-CsTenant -ErrorAction SilentlyContinue).DisplayName
        }
        DataSummary = @{}
    }
    
    foreach ($key in $CollectedData.Keys) {
        if ($CollectedData[$key] -is [array]) {
            $summary.DataSummary[$key] = @{
                Count = $CollectedData[$key].Count
                HasData = $CollectedData[$key].Count -gt 0
            }
        }
        elseif ($CollectedData[$key] -is [hashtable]) {
            $summary.DataSummary[$key] = @{
                Keys = @($CollectedData[$key].Keys)
                HasData = $CollectedData[$key].Keys.Count -gt 0
            }
        }
    }
    
    Export-DataToJson -Data $summary -OutputPath $OutputPath -FileName "00_Summary"
}

# Function to test for call flow generation prerequisites
function Test-CallFlowPrerequisites {
    $callFlowScript = Join-Path $ModulesPath "Generate-CallFlowMaps-Simple.ps1"
    $pdfGenerator = Join-Path $ModulesPath "pdf_generator.py"
    
    $prerequisites = @{
        CallFlowScript = Test-Path $callFlowScript
        PdfGenerator = Test-Path $pdfGenerator
        Python = $false
        Playwright = $false
    }
    
    # Test for Python
    try {
        $pythonCmd = if ($IsWindows) { "python" } else { "python3" }
        $pythonVersion = & $pythonCmd --version 2>&1
        if ($pythonVersion -match "Python \d+\.\d+") {
            $prerequisites.Python = $true
        }
    }
    catch {
        # Python not found
        $prerequisites.Python = $false
    }
    
    # Test for Playwright (only if Python is available)
    if ($prerequisites.Python) {
        try {
            $pythonCmd = if ($IsWindows) { "python" } else { "python3" }
            $playwrightTest = & $pythonCmd -c "import playwright; print('OK')" 2>&1
            if ($playwrightTest -match "OK") {
                $prerequisites.Playwright = $true
            }
        }
        catch {
            $prerequisites.Playwright = $false
        }
    }
    else {
        $prerequisites.Playwright = $false
    }
    
    return $prerequisites
}

# Function to generate call flow maps
function Invoke-CallFlowGeneration {
    param(
        [string]$DataPath,
        [bool]$GeneratePDF = $false,
        [string[]]$FilterNumbers = @()
    )
    
    Write-Host "`n=== Generating Call Flow Maps ===" -ForegroundColor Cyan
    
    $callFlowScript = Join-Path $ModulesPath "Generate-CallFlowMaps-Simple.ps1"
    
    if (-not (Test-Path $callFlowScript)) {
        Write-Warning "Call flow generation script not found: $callFlowScript"
        return $false
    }
    
    # Test prerequisites
    $prereqs = Test-CallFlowPrerequisites
    
    Write-Host "Prerequisites check:" -ForegroundColor Gray
    Write-Host "  ✓ Call flow script: $($prereqs.CallFlowScript)" -ForegroundColor $(if ($prereqs.CallFlowScript) { "Green" } else { "Red" })
    Write-Host "  ✓ PDF generator: $($prereqs.PdfGenerator)" -ForegroundColor $(if ($prereqs.PdfGenerator) { "Green" } else { "Red" })
    Write-Host "  ✓ Python available: $($prereqs.Python)" -ForegroundColor $(if ($prereqs.Python) { "Green" } else { "Yellow" })
    Write-Host "  ✓ Playwright available: $($prereqs.Playwright)" -ForegroundColor $(if ($prereqs.Playwright) { "Green" } else { "Yellow" })
    
    if (-not $prereqs.CallFlowScript) {
        Write-Error "Cannot generate call flow maps - script missing"
        return $false
    }
    
    # Build arguments for call flow script
    $callFlowArgs = @()
    $callFlowArgs += "-JsonDataPath"
    $callFlowArgs += "`"$DataPath`""
    
    if ($FilterNumbers.Count -gt 0) {
        $callFlowArgs += "-FilterByNumber"
        foreach ($number in $FilterNumbers) {
            $callFlowArgs += "`"$number`""
        }
    }
    
    # Add IncludeDetailedSettings for comprehensive output
    $callFlowArgs += "-IncludeDetailedSettings"
    
    try {
        # Generate HTML call flows
        Write-Host "Generating HTML call flow maps..." -ForegroundColor Yellow
        
        # Construct the full argument list properly
        $allArgs = @("-NoProfile", "-File", "`"$callFlowScript`"")
        $allArgs += $callFlowArgs
        
        $process = Start-Process -FilePath "pwsh" -ArgumentList $allArgs -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "✓ HTML call flow maps generated successfully" -ForegroundColor Green
            
            # Generate PDFs if requested and prerequisites are met
            if ($GeneratePDF) {
                if ($prereqs.Python -and $prereqs.Playwright) {
                    Write-Host "Generating PDF files..." -ForegroundColor Yellow
                    
                    # Find the most recent call flow output directory
                    $callFlowDirs = Get-ChildItem -Path $PSScriptRoot -Directory -Name "CallFlowMaps_*" | Sort-Object -Descending
                    if ($callFlowDirs.Count -gt 0) {
                        $latestCallFlowDir = Join-Path $PSScriptRoot $callFlowDirs[0]
                        $individualDir = Join-Path $latestCallFlowDir "Individual"
                        $summaryDir = Join-Path $latestCallFlowDir "Summary"
                        $pdfDir = Join-Path $latestCallFlowDir "PDF"
                        
                        # Generate PDFs using Python script
                        $pdfGenerator = Join-Path $ModulesPath "pdf_generator.py"
                        
                        # Determine Python command
                        $pythonCmd = if ($IsWindows) { "python" } else { "python3" }
                        
                        # Convert Individual HTML files
                        if (Test-Path $individualDir) {
                            $individualPdfDir = Join-Path $pdfDir "Individual"
                            $pdfProcess = Start-Process -FilePath $pythonCmd -ArgumentList "`"$pdfGenerator`"", "`"$individualDir`"", "-o", "`"$individualPdfDir`"" -Wait -NoNewWindow -PassThru
                            
                            if ($pdfProcess.ExitCode -eq 0) {
                                Write-Host "✓ Individual call flow PDFs generated" -ForegroundColor Green
                            } else {
                                Write-Warning "Individual PDF generation completed with warnings"
                            }
                        }
                        
                        # Convert Summary HTML files
                        if (Test-Path $summaryDir) {
                            $summaryPdfDir = Join-Path $pdfDir "Summary"
                            $pdfProcess = Start-Process -FilePath $pythonCmd -ArgumentList "`"$pdfGenerator`"", "`"$summaryDir`"", "-o", "`"$summaryPdfDir`"" -Wait -NoNewWindow -PassThru
                            
                            if ($pdfProcess.ExitCode -eq 0) {
                                Write-Host "✓ Summary dashboard PDF generated" -ForegroundColor Green
                            } else {
                                Write-Warning "Summary PDF generation completed with warnings"
                            }
                        }
                        
                        Write-Host "✓ PDF generation complete" -ForegroundColor Green
                    } else {
                        Write-Warning "Could not find call flow output directory for PDF generation"
                    }
                } else {
                    Write-Warning "PDF generation requested but prerequisites not met:"
                    if (-not $prereqs.Python) {
                        Write-Warning "  - Python not available. Install Python 3.7+ to enable PDF generation"
                    }
                    if (-not $prereqs.Playwright) {
                        Write-Warning "  - Playwright not available. Install with: pip install playwright"
                        Write-Warning "  - Then install browsers with: playwright install"
                    }
                    if (-not $prereqs.PdfGenerator) {
                        Write-Warning "  - PDF generator script not found"
                    }
                }
            }
            
            return $true
        } else {
            Write-Warning "Call flow generation completed with exit code: $($process.ExitCode)"
            return $false
        }
    }
    catch {
        Write-Error "Failed to generate call flow maps: $($_.Exception.Message)"
        return $false
    }
}

# Function to create summary report
function New-SummaryReport {
    param(
        [hashtable]$CollectedData,
        [string]$OutputPath
    )
    
    $summary = @{
        CollectionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
        TenantInfo = @{
            TenantId = (Get-CsTenant -ErrorAction SilentlyContinue).TenantId
            DisplayName = (Get-CsTenant -ErrorAction SilentlyContinue).DisplayName
        }
        DataSummary = @{}
    }
    
    foreach ($key in $CollectedData.Keys) {
        if ($CollectedData[$key] -is [array]) {
            $summary.DataSummary[$key] = @{
                Count = $CollectedData[$key].Count
                HasData = $CollectedData[$key].Count -gt 0
            }
        }
        elseif ($CollectedData[$key] -is [hashtable]) {
            $summary.DataSummary[$key] = @{
                Keys = @($CollectedData[$key].Keys)
                HasData = $CollectedData[$key].Keys.Count -gt 0
            }
        }
    }
    
    Export-DataToJson -Data $summary -OutputPath $OutputPath -FileName "00_Summary"
}

# Main execution
try {
    Write-Host "=== Microsoft Teams Calling Data Collection ===" -ForegroundColor Cyan
    Write-Host "Started at: $(Get-Date)" -ForegroundColor Gray
    
    # Initialize output directory
    $outputDir = Initialize-OutputDirectory -Path $OutputPath
    Write-Host "Output directory: $outputDir" -ForegroundColor Gray
    
    # Check for required module
    if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
        Write-Error "MicrosoftTeams module is not installed. Please install it using: Install-Module -Name MicrosoftTeams"
        exit 1
    }
    
    # Connect to Teams if not already connected
    if (-not (Test-TeamsConnection)) {
        if (-not (Connect-TeamsService -TenantId $TenantId)) {
            Write-Error "Failed to connect to Microsoft Teams. Exiting."
            exit 1
        }
    }
    
    # Load custom modules
    Write-Host "`nLoading custom modules..." -ForegroundColor Cyan
    $loadedModules = Import-CustomModules -Modules $ModuleList -ModulesPath $ModulesPath
    
    if ($loadedModules.Count -eq 0) {
        Write-Warning "No custom modules loaded. Some functionality may be limited."
    }
    
    # Initialize data collection hashtable
    $allData = @{}
    
    # Collect data from each module
    Write-Host "`nCollecting Teams calling data..." -ForegroundColor Cyan
    
    # 1. Calling Policies
    if ("TeamsCallingPolicies" -in $loadedModules) {
        Write-Host "Gathering calling policies..." -ForegroundColor Yellow
        $allData.CallingPolicies = Get-TeamsCallingPoliciesData
        Export-DataToJson -Data $allData.CallingPolicies -OutputPath $outputDir -FileName "01_CallingPolicies"
    }
    
    # 2. Call Queues
    if ("TeamsCallQueues" -in $loadedModules) {
        Write-Host "Gathering call queues..." -ForegroundColor Yellow
        $allData.CallQueues = Get-TeamsCallQueuesData
        Export-DataToJson -Data $allData.CallQueues -OutputPath $outputDir -FileName "02_CallQueues"
    }
    
    # 3. Auto Attendants
    if ("TeamsAutoAttendants" -in $loadedModules) {
        Write-Host "Gathering auto attendants..." -ForegroundColor Yellow
        $allData.AutoAttendants = Get-TeamsAutoAttendantsData
        Export-DataToJson -Data $allData.AutoAttendants -OutputPath $outputDir -FileName "03_AutoAttendants"
    }
    
    # 4. User Settings
    if ("TeamsUserSettings" -in $loadedModules) {
        Write-Host "Gathering user settings..." -ForegroundColor Yellow
        $allData.UserSettings = Get-TeamsUserSettingsData -IncludeDetailedData:$IncludeUserData
        Export-DataToJson -Data $allData.UserSettings -OutputPath $outputDir -FileName "04_UserSettings"
    }
    
    # 5. Emergency Locations
    if ("TeamsEmergencyLocations" -in $loadedModules) {
        Write-Host "Gathering emergency locations..." -ForegroundColor Yellow
        $allData.EmergencySettings = Get-TeamsEmergencyLocationsData
        Export-DataToJson -Data $allData.EmergencySettings -OutputPath $outputDir -FileName "05_EmergencySettings"
    }
    
    # 6. Voice Routing
    if ("TeamsVoiceRouting" -in $loadedModules) {
        Write-Host "Gathering voice routing..." -ForegroundColor Yellow
        $allData.VoiceRouting = Get-TeamsVoiceRoutingData
        Export-DataToJson -Data $allData.VoiceRouting -OutputPath $outputDir -FileName "06_VoiceRouting"
    }
    
    # 7. Compliance Settings
    if ("TeamsComplianceSettings" -in $loadedModules) {
        Write-Host "Gathering compliance settings..." -ForegroundColor Yellow
        $allData.ComplianceSettings = Get-TeamsComplianceSettingsData
        Export-DataToJson -Data $allData.ComplianceSettings -OutputPath $outputDir -FileName "07_ComplianceSettings"
    }
    
    # Create summary report
    Write-Host "Creating summary report..." -ForegroundColor Yellow
    New-SummaryReport -CollectedData $allData -OutputPath $outputDir
    
    # Generate call flow maps if requested
    if ($GenerateCallFlowMaps) {
        $callFlowSuccess = Invoke-CallFlowGeneration -DataPath $outputDir -GeneratePDF:$GeneratePDF -FilterNumbers $CallFlowFilterNumbers
        
        if ($callFlowSuccess) {
            Write-Host "✓ Call flow maps generated successfully" -ForegroundColor Green
            
            # Find the most recent call flow directory for reporting
            $callFlowDirs = Get-ChildItem -Path $PSScriptRoot -Directory -Name "CallFlowMaps_*" | Sort-Object -Descending
            if ($callFlowDirs.Count -gt 0) {
                $latestCallFlowDir = Join-Path $PSScriptRoot $callFlowDirs[0]
                $dashboardFile = Join-Path $latestCallFlowDir "Summary" "Dashboard.html"
                
                if (Test-Path $dashboardFile) {
                    Write-Host "Call flow dashboard: $dashboardFile" -ForegroundColor Cyan
                }
            }
        } else {
            Write-Warning "Call flow map generation encountered issues"
        }
    }
    
    # Compress output if requested
    if ($CompressOutput) {
        Write-Host "Compressing output..." -ForegroundColor Yellow
        $zipPath = "$outputDir.zip"
        try {
            Compress-Archive -Path "$outputDir\*" -DestinationPath $zipPath -Force
            Write-Host "✓ Created compressed archive: $zipPath" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to create compressed archive: $($_.Exception.Message)"
        }
    }
    
    # Final summary
    Write-Host "`n=== Collection Complete ===" -ForegroundColor Green
    Write-Host "Completed at: $(Get-Date)" -ForegroundColor Gray
    Write-Host "Output location: $outputDir" -ForegroundColor Gray
    
    $totalFiles = (Get-ChildItem -Path $outputDir -Filter "*.json").Count
    Write-Host "Total JSON files created: $totalFiles" -ForegroundColor Gray
    
    if ($GenerateCallFlowMaps) {
        $callFlowDirs = Get-ChildItem -Path $PSScriptRoot -Directory -Name "CallFlowMaps_*" | Sort-Object -Descending
        if ($callFlowDirs.Count -gt 0) {
            $latestCallFlowDir = Join-Path $PSScriptRoot $callFlowDirs[0]
            Write-Host "Call flow maps location: $latestCallFlowDir" -ForegroundColor Gray
            
            if ($GeneratePDF) {
                $pdfDir = Join-Path $latestCallFlowDir "PDF"
                if (Test-Path $pdfDir) {
                    $pdfCount = (Get-ChildItem -Path $pdfDir -Filter "*.pdf" -Recurse).Count
                    Write-Host "PDF files generated: $pdfCount" -ForegroundColor Gray
                }
            }
        }
    }
    
    if ($CompressOutput -and (Test-Path "$outputDir.zip")) {
        $zipSize = [math]::Round((Get-Item "$outputDir.zip").Length / 1MB, 2)
        Write-Host "Compressed archive size: $zipSize MB" -ForegroundColor Gray
    }
}
catch {
    Write-Error "An error occurred during execution: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
}
finally {
    # Cleanup
    Write-Host "`nCleaning up..." -ForegroundColor Gray
    
    # Remove imported custom modules
    foreach ($moduleName in $loadedModules) {
        try {
            Remove-Module $moduleName -Force -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore cleanup errors
        }
    }
}

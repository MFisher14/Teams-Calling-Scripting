# Teams Calling Data Collection Tool

This PowerShell script system provides comprehensive collection of Microsoft Teams calling data, exporting all settings and configurations to JSON files for analysis and troubleshooting.

## Overview

The tool consists of a master script (`Gather-Teams-Calling-Data.ps1`) that orchestrates data collection using specialized modules for different Teams calling components.

## Prerequisites

### Required PowerShell Modules
```powershell
# Install the Microsoft Teams PowerShell module
Install-Module -Name MicrosoftTeams -Force -Scope CurrentUser

# Optional: For enhanced compliance data collection
Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
```

### Required Permissions
The account running this script needs:
- **Teams Administrator** or **Global Administrator** role
- **Teams Communications Administrator** (for calling policies and configurations)
- **Teams Communications Support Engineer** (for call analytics data)
- **Compliance Administrator** (for compliance and security settings)

## Components

### Master Script
- `Gather-Teams-Calling-Data.ps1` - Main orchestration script

### Modules (`/Modules/`)
1. **TeamsCallingPolicies.psm1** - Calling policies and configurations
2. **TeamsCallQueues.psm1** - Call queues and hunt groups
3. **TeamsAutoAttendants.psm1** - Auto attendants and schedules
4. **TeamsUserSettings.psm1** - User calling settings and policy assignments
5. **TeamsEmergencyLocations.psm1** - Emergency calling locations and policies
6. **TeamsVoiceRouting.psm1** - Voice routing, PSTN, and Direct Routing
7. **TeamsComplianceSettings.psm1** - Compliance, security, and audit settings

## Usage

### Basic Usage
```powershell
# Run with default settings (creates timestamped folder in current directory)
.\Gather-Teams-Calling-Data.ps1
```

### Advanced Usage
```powershell
# Full data collection with user details
.\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\TeamsData" -IncludeUserData -CompressOutput

# Specify tenant ID for authentication
.\Gather-Teams-Calling-Data.ps1 -TenantId "your-tenant-id-here" -IncludeUserData
```

### Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `OutputPath` | String | Directory where JSON files will be saved | Current directory with timestamp |
| `TenantId` | String | Specific tenant ID for authentication | Uses interactive login |
| `IncludeUserData` | Switch | Include detailed per-user calling data | False (summary only) |
| `CompressOutput` | Switch | Create ZIP archive of all JSON files | False |

## Output Structure

The script creates JSON files for each component:

```
TeamsCallingData_20250926_143022/
├── 00_Summary.json                    # Overview and collection summary
├── 01_CallingPolicies.json           # All calling policies and configurations
├── 02_CallQueues.json                # Call queues, hunt groups, and agents
├── 03_AutoAttendants.json            # Auto attendants, schedules, and flows
├── 04_UserSettings.json              # User calling settings and assignments
├── 05_EmergencySettings.json         # Emergency locations and policies
├── 06_VoiceRouting.json              # Voice routes, SBCs, and PSTN settings
└── 07_ComplianceSettings.json        # Compliance, security, and audit settings
```

## Data Collected

### 1. Calling Policies (`01_CallingPolicies.json`)
- Teams calling policies
- Voice application policies
- IP phone policies
- Branch survivability policies
- Dialout policies
- Network configuration
- Calling line identity policies

### 2. Call Queues (`02_CallQueues.json`)
- Call queue configurations
- Agent assignments and distribution lists
- Overflow and timeout actions
- Resource account associations
- Queue statistics and permissions

### 3. Auto Attendants (`03_AutoAttendants.json`)
- Auto attendant configurations
- Call flows and menu structures
- Schedules (business hours and holidays)
- Language and voice settings
- Resource account associations

### 4. User Settings (`04_UserSettings.json`)
- Voice-enabled user accounts
- Policy assignments per user
- Line URI assignments
- Calling plan vs. Direct Routing assignments
- Location assignments
- Policy distribution summary

### 5. Emergency Settings (`05_EmergencySettings.json`)
- Emergency calling policies
- Emergency call routing policies
- Emergency addresses and locations
- Network sites and trusted IPs
- Location Information Service (LIS) configuration
- User emergency policy assignments

### 6. Voice Routing (`06_VoiceRouting.json`)
- Voice routing policies and routes
- PSTN usages and gateways
- Session Border Controller (SBC) configurations
- Tenant dial plans and normalization rules
- Phone number assignments
- Number porting information

### 7. Compliance Settings (`07_ComplianceSettings.json`)
- Compliance recording policies
- Call recording policies
- Information barrier policies
- Security and privacy settings
- Audit configurations
- Data retention policies
- Call analytics settings

## Performance Considerations

### Large Tenants (1000+ Users)
- Use `-IncludeUserData` cautiously - it significantly increases collection time
- Consider running during off-peak hours
- Monitor script progress through console output

### Network Considerations
- Script makes multiple API calls to Microsoft Graph and Teams services
- Ensure stable internet connection
- May take 15-60 minutes depending on tenant size and data scope

## Troubleshooting

### Common Issues

#### Authentication Failures
```
Error: Failed to connect to Microsoft Teams
Solution: Ensure you have appropriate admin permissions and MFA is configured
```

#### Module Loading Errors
```
Warning: Module not found: TeamsCallingPolicies.psm1
Solution: Ensure all module files are in the /Modules/ subdirectory
```

#### Permission Denied Errors
```
Error: Access denied when collecting [specific data]
Solution: Verify admin role assignments and API permissions
```

### Debug Mode
Enable verbose output for troubleshooting:
```powershell
$VerbosePreference = "Continue"
.\Gather-Teams-Calling-Data.ps1 -Verbose
```

## Security Considerations

### Data Sensitivity
- Output files contain sensitive configuration data
- Phone numbers, user identities, and policy assignments
- Store output files securely and restrict access

### Credential Management
- Script uses interactive authentication by default
- Credentials are not stored or cached
- Consider using service accounts for automated runs

## Examples

### Scenario 1: Full Configuration Audit
```powershell
# Comprehensive audit including all user data
.\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\Audit\Teams_$(Get-Date -Format 'yyyy-MM-dd')" -IncludeUserData -CompressOutput
```

### Scenario 2: Quick Policy Review
```powershell
# Fast collection of policies and configurations (no detailed user data)
.\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\Temp\QuickScan"
```

### Scenario 3: Specific Tenant Analysis
```powershell
# Target specific tenant with full data collection
.\Gather-Teams-Calling-Data.ps1 -TenantId "contoso.onmicrosoft.com" -IncludeUserData
```

## Integration

### PowerBI Analysis
The JSON output can be imported into Power BI for visualization:
1. Use Power BI's JSON connector
2. Import the summary file first for overview
3. Create relationships between datasets using common identifiers

### Excel Analysis
Convert JSON to Excel for detailed analysis:
```powershell
# Example: Convert summary to CSV
Get-Content "00_Summary.json" | ConvertFrom-Json | ConvertTo-Csv -NoTypeInformation | Out-File "Summary.csv"
```

### Automated Monitoring
Schedule regular collection for change tracking:
```powershell
# Example scheduled task
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 6:00AM
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\Gather-Teams-Calling-Data.ps1 -OutputPath C:\MonthlyReports -CompressOutput"
Register-ScheduledTask -TaskName "Teams Calling Data Collection" -Trigger $Trigger -Action $Action
```

## Support

### Log Files
The script outputs detailed progress information to the console. Redirect to a file for logging:
```powershell
.\Gather-Teams-Calling-Data.ps1 -IncludeUserData *> "collection_log.txt"
```

### Version Information
- Script Version: 1.0
- Compatible with: PowerShell 5.1+ and PowerShell 7+
- Microsoft Teams Module: 4.0.0+
- Last Updated: September 26, 2025

---

*This tool is designed for TCP Calling Issues analysis and Teams calling configuration auditing.*
# Microsoft Teams Calling Data Collection & Call Flow Analysis Suite

This comprehensive PowerShell solution collects all Microsoft Teams calling configuration data and generates visual call flow maps with PDF export capability. Perfect for Teams administrators, consultants, and support engineers who need to understand and document complex calling configurations.

## 🌟 Key Features

- **📊 Complete Data Collection**: Gathers all Teams calling settings, policies, phone numbers, auto attendants, call queues, emergency locations, and more
- **🧩 Modular Architecture**: 7 specialized modules for different Teams components plus call flow generators
- **🎨 Visual Call Flow Maps**: Generates professional HTML call flow diagrams for each phone number
- **📄 PDF Export**: Converts call flow maps to PDF using modern Playwright technology
- **📋 Summary Dashboard**: Interactive overview of all phone numbers and their configurations
- **🎯 Flexible Filtering**: Focus on specific phone numbers or collect all data
- **⚡ Automated Workflow**: Seamless integration from data collection to PDF generation
- **🔍 Advanced Analysis**: Step-by-step visual call routing with detailed configuration data
- **👥 Queue User Details**: Shows which users are assigned to each call queue
- **🎯 Flexible Filtering**: Focus on specific phone numbers or collect all data
- **⚡ Automated Workflow**: Seamless integration from data collection to PDF generation
- **🔍 Advanced Analysis**: Step-by-step visual call routing with detailed configuration data

## 🚀 Quick Start

### 1. Install Prerequisites
```bash
# Install Python dependencies for PDF generation (choose one method)
python Modules/setup_pdf.py          # Automated setup
# OR manual install:
pip install playwright && playwright install chromium
```

### 2. Run Complete Analysis
```powershell
# Basic data collection only
.\Gather-Teams-Calling-Data.ps1

# Full analysis with visual call flow maps and PDFs
.\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps -GeneratePDF

# Filter to specific phone numbers for focused analysis
.\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps -GeneratePDF -CallFlowFilterNumbers "+12125551234", "+13105556789"
```

## 📋 System Requirements

### Required
- **PowerShell**: 5.1 or newer (PowerShell 7 recommended)
- **Microsoft Teams Module**: `Install-Module MicrosoftTeams`
- **Administrative Permissions**: Teams Administrator or Global Administrator role

### Optional (for PDF Generation)
- **Python**: 3.7 or higher 
- **Playwright**: Modern browser automation (replaces discontinued wkhtmltopdf)

### Additional Permissions for Complete Data Collection
- **Teams Communications Administrator** (for calling policies and configurations)
- **Teams Communications Support Engineer** (for call analytics data)
- **Compliance Administrator** (for compliance and security settings)

## 📁 Project Structure

```
TMC Calling Issues/
├── README.md                           # This file - complete documentation
├── Gather-Teams-Calling-Data.ps1      # Main orchestration script
├── Initial_Issue_Log.md                # Project background and issues
├── Modules/                            # All modules and components
│   ├── TeamsCallingPolicies.psm1      # Calling policies and configurations
│   ├── TeamsCallQueues.psm1           # Call queues and hunt groups  
│   ├── TeamsAutoAttendants.psm1       # Auto attendants and schedules
│   ├── TeamsUserSettings.psm1         # User calling settings and assignments
│   ├── TeamsEmergencyLocations.psm1   # Emergency calling locations and policies
│   ├── TeamsVoiceRouting.psm1         # Voice routing, PSTN, and Direct Routing
│   ├── TeamsComplianceSettings.psm1   # Compliance, security, and audit settings
│   ├── Generate-CallFlowMaps-Simple.ps1 # Modern call flow map generator
│   ├── Generate-CallFlowMaps.ps1       # Legacy call flow generator (PSWriteHTML)
│   ├── pdf_generator.py                # Python PDF generation using Playwright
│   └── setup_pdf.py                    # Automated Python environment setup
└── [Generated Output Folders]
    ├── TeamsCallingData_YYYYMMDD_HHMMSS/  # JSON data files
    └── CallFlowMaps_YYYYMMDD_HHMMSS/      # HTML and PDF call flow maps
```

## 🎯 Core Components

### Master Script
- **`Gather-Teams-Calling-Data.ps1`** - Main orchestration script that collects all data and optionally generates call flow maps

### Data Collection Modules (`/Modules/`)
1. **`TeamsCallingPolicies.psm1`** - Calling policies and configurations
2. **`TeamsCallQueues.psm1`** - Call queues and hunt groups
3. **`TeamsAutoAttendants.psm1`** - Auto attendants and schedules  
4. **`TeamsUserSettings.psm1`** - User calling settings and policy assignments
5. **`TeamsEmergencyLocations.psm1`** - Emergency calling locations and policies
6. **`TeamsVoiceRouting.psm1`** - Voice routing, PSTN, and Direct Routing
7. **`TeamsComplianceSettings.psm1`** - Compliance, security, and audit settings

### Call Flow Analysis Tools (`/Modules/`)
- **`Generate-CallFlowMaps-Simple.ps1`** - Modern HTML call flow generator with clean design
- **`Generate-CallFlowMaps.ps1`** - Legacy PSWriteHTML-based generator (deprecated)
- **`pdf_generator.py`** - Python-based PDF conversion using Playwright
- **`setup_pdf.py`** - Automated setup for Python PDF dependencies

## 💼 Usage Examples

### Basic Data Collection
```powershell
# Simple collection - creates timestamped folder in current directory
.\Gather-Teams-Calling-Data.ps1

# Specify custom output path
.\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\TeamsAudit\2025-09-26"

# Include detailed user data (takes longer for large tenants)
.\Gather-Teams-Calling-Data.ps1 -IncludeUserData -CompressOutput
```

### Call Flow Analysis
```powershell
# Generate HTML call flow maps for all phone numbers  
.\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps

# Full analysis with PDF export
.\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps -GeneratePDF

# Focus on specific numbers (performance optimization)
.\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps -GeneratePDF -CallFlowFilterNumbers "+1984308", "+1800"
```

### Advanced Scenarios
```powershell
# Complete audit for specific tenant
.\Gather-Teams-Calling-Data.ps1 -TenantId "contoso.onmicrosoft.com" -IncludeUserData -GenerateCallFlowMaps -GeneratePDF -CompressOutput

# Monthly automated reporting
.\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\Reports\Teams_$(Get-Date -Format 'yyyy-MM')" -IncludeUserData -GenerateCallFlowMaps -GeneratePDF -CompressOutput
```

## 📊 Parameters Reference

### Main Script Parameters (`Gather-Teams-Calling-Data.ps1`)

| Parameter | Type | Description | Default | Example |
|-----------|------|-------------|---------|---------|
| `OutputPath` | String | Directory for JSON files | Current directory + timestamp | `"C:\TeamsData"` |
| `TenantId` | String | Specific tenant ID for auth | Interactive login | `"contoso.onmicrosoft.com"` |
| `IncludeUserData` | Switch | Include detailed user data | False | `-IncludeUserData` |
| `CompressOutput` | Switch | Create ZIP archive | False | `-CompressOutput` |
| `GenerateCallFlowMaps` | Switch | Generate HTML call flows | False | `-GenerateCallFlowMaps` |
| `GeneratePDF` | Switch | Generate PDF files | False | `-GeneratePDF` |
| `CallFlowFilterNumbers` | String[] | Filter to specific numbers | All numbers | `"+1984308", "+1800"` |

### Call Flow Generator Parameters

| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `JsonDataPath` | String | Path to JSON data directory | `".\TeamsCallingData_20250926_110848"` |
| `OutputPath` | String | Custom output directory | `"C:\CallFlowMaps"` |
| `GeneratePDF` | Switch | Generate PDF files | `-GeneratePDF` |
| `IncludeDetailedSettings` | Switch | Include detailed configuration | `-IncludeDetailedSettings` |
| `FilterByNumber` | String[] | Filter to specific numbers | `"+1984308", "+1800"` |

## 📈 Output Structure

### JSON Data Files
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

### Call Flow Maps
```
CallFlowMaps_20250926_112010/
├── Summary/
│   └── Dashboard.html                 # Interactive summary dashboard
├── Individual/
│   ├── 19843081541.html              # Call flow for +1-984-308-1541
│   ├── 19843081461.html              # Call flow for +1-984-308-1461  
│   └── [other numbers].html
├── PDF/ (if -GeneratePDF used)
│   ├── Summary/
│   │   └── Dashboard.pdf
│   └── Individual/
│       ├── 19843081541.pdf
│       └── [other number PDFs]
└── HTML/ (backup)
```

## 🔍 Data Collection Details

### 1. Calling Policies (`01_CallingPolicies.json`)
- Teams calling policies and restrictions
- Voice application policies  
- IP phone policies and device settings
- Branch survivability policies
- Dialout policies and restrictions
- Network configuration and sites
- Calling line identity policies
- Privacy configuration

### 2. Call Queues (`02_CallQueues.json`)
- Call queue configurations and settings
- Agent assignments and distribution lists
- Overflow and timeout actions
- Resource account associations
- Queue statistics and permissions
- Hunt group configurations

### 3. Auto Attendants (`03_AutoAttendants.json`)
- Auto attendant configurations
- Call flows and menu structures  
- Business hours and holiday schedules
- Language and voice settings
- Resource account associations
- Call handling associations

### 4. User Settings (`04_UserSettings.json`)
- Voice-enabled user accounts
- Policy assignments per user
- Line URI and phone number assignments
- Calling plan vs. Direct Routing assignments
- Emergency location assignments
- Policy distribution summary and analytics

### 5. Emergency Settings (`05_EmergencySettings.json`)
- Emergency calling policies
- Emergency call routing policies
- Emergency addresses and locations
- Network sites and trusted IP addresses
- Location Information Service (LIS) configuration
- User emergency policy assignments

### 6. Voice Routing (`06_VoiceRouting.json`)
- Voice routing policies and routes
- PSTN usages and route priorities
- Session Border Controller (SBC) configurations
- Tenant dial plans and normalization rules
- Phone number assignments and capabilities
- Number porting information and status

### 7. Compliance Settings (`07_ComplianceSettings.json`)
- Compliance recording policies
- Call recording policies and applications
- Information barrier policies
- Security and privacy settings
- Audit configurations and logging
- Data retention policies
- Call analytics settings and permissions

## 🎨 Call Flow Analysis Features

### Visual Call Flow Types

#### 👤 User-Assigned Numbers
**Example Flow**: 
1. 📞 **Incoming Call** - PSTN Gateway receives call
2. 👤 **User Routing** - Route directly to assigned user  
3. 🔧 **Voice Policy** - Apply calling policies (if configured)

**Shows**: User details, Enterprise Voice status, voice routing policies

#### 🏢 Call Queue Numbers  
**Example Flow**:
1. 📞 **Incoming Call** - Resource account receives call
2. 🏢 **Call Queue Processing** - Enter call queue system
3. ⚙️ **Queue Settings** - Apply routing method and settings
4. ⚠️ **Overflow Action** - Handle queue overflow (if configured)
5. ⏱️ **Timeout Action** - Handle call timeouts (if configured)

**Shows**: Queue name, routing method, agent assignments, overflow/timeout settings

#### 🤖 Auto Attendant Numbers
**Example Flow**:
1. 📞 **Incoming Call** - Resource account receives call
2. 🤖 **Auto Attendant Processing** - Enter auto attendant system
3. 🗣️ **Language & Voice** - Configure language and voice settings  
4. 📋 **Menu Presentation** - Present menu options to caller
5. 📅 **Call Handling** - Apply business hours and holiday routing

**Shows**: Attendant name, language settings, menu options, schedules

### 🎨 Visual Design Features
- **Color-coded components** for easy identification
- **Interactive hover effects** and responsive design  
- **Emoji icons** for quick visual reference
- **Step-by-step flow diagrams** with clear progression
- **Expandable sections** for detailed configuration data
- **Professional styling** suitable for client presentations

## 🛠️ PDF Generation

### Modern Playwright Approach
The suite uses Python with Playwright for reliable, high-quality PDF generation, replacing the discontinued wkhtmltopdf tool.

### Setup PDF Generation
```bash
# Automated setup (recommended)
python Modules/setup_pdf.py

# Manual setup  
pip install playwright
playwright install chromium
```

### PDF Features
- **High-quality rendering** with preserved formatting
- **Embedded fonts** and styling
- **Print-optimized layouts**
- **Individual PDFs** for each phone number
- **Summary dashboard PDF**
- **Cross-platform compatibility** (Windows, macOS, Linux)

## ⚡ Performance Optimization

### Large Tenant Considerations (1000+ Users)
- **Use filtering**: `-CallFlowFilterNumbers` for targeted analysis
- **Avoid detailed user data**: Only use `-IncludeUserData` when necessary
- **Run during off-peak hours**: API rate limiting consideration
- **Monitor progress**: Detailed console output shows progress

### Processing Time Estimates
- **Basic data collection**: 5-15 minutes (depends on tenant size)
- **With user data**: 15-60 minutes for large tenants  
- **Call flow generation**: 1-2 minutes per 100 numbers
- **PDF generation**: 30-60 seconds per 100 numbers

### Memory Usage
- **Small tenants (<100 numbers)**: ~200 MB
- **Medium tenants (100-500 numbers)**: ~500 MB  
- **Large tenants (1000+ numbers)**: ~1-2 GB

## 🔧 Troubleshooting

### Common Issues and Solutions

#### Authentication Problems
```
Error: Failed to connect to Microsoft Teams
Solutions:
- Verify admin permissions (Teams Administrator or Global Administrator)  
- Check MFA configuration and complete authentication
- Try specifying tenant ID: -TenantId "yourtenant.onmicrosoft.com"
```

#### Module Loading Issues
```
Warning: Module not found: TeamsCallingPolicies.psm1
Solutions:
- Ensure all .psm1 files are in the /Modules/ subdirectory
- Check file permissions and unblock if downloaded from internet
- Verify PowerShell execution policy allows module loading
```

#### No Phone Numbers Found
```
Issue: ✓ Found 0 phone numbers  
Solutions:
- Verify phone numbers are assigned in Teams Admin Center
- Check that users have voice licenses (Calling Plan or Phone System)
- Ensure resource accounts have phone numbers assigned
- Review VoiceRouting.json for PhoneNumberAssignments data
```

#### PDF Generation Failures
```
Issue: PDF generation requested but prerequisites not met
Solutions:
- Install Python 3.7+: Download from python.org
- Install Playwright: pip install playwright && playwright install chromium  
- Run setup script: python Modules/setup_pdf.py
- Check Python command availability: python3 --version (macOS) or python --version (Windows)
```

#### Performance Issues
```
Issue: Script runs slowly or times out
Solutions:
- Use -CallFlowFilterNumbers to focus on specific numbers
- Avoid -IncludeUserData for large tenants unless necessary
- Run during off-peak hours to avoid API throttling
- Consider breaking large tenants into smaller filtered runs
```

### Debug Mode
Enable detailed logging for troubleshooting:
```powershell
$VerbosePreference = "Continue"  
.\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps -GeneratePDF -Verbose
```

### Log Output
```powershell
# Capture all output to file for analysis
.\Gather-Teams-Calling-Data.ps1 -GenerateCallFlowMaps -GeneratePDF *> "collection_log.txt"
```

## 🔒 Security Considerations

### Data Sensitivity
- **Highly sensitive data**: Phone numbers, user identities, policy configurations
- **Secure storage required**: Store output files in protected locations
- **Access control**: Restrict access to authorized personnel only
- **Data retention**: Follow organizational data retention policies

### Credential Management  
- **Interactive authentication**: Script uses browser-based login (no stored credentials)
- **Service accounts**: Consider dedicated service accounts for automated runs
- **MFA compliance**: Supports multi-factor authentication requirements
- **Token management**: Tokens are session-based and not persisted

### Best Practices
```powershell
# Run with minimal permissions needed
.\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\SecureLocation\TeamsAudit"

# Immediately compress and secure output
.\Gather-Teams-Calling-Data.ps1 -CompressOutput
# Then move ZIP to secure archive location
```

## 📈 Integration and Analysis

### PowerBI Integration
```powershell
# Convert JSON to PowerBI-friendly format
$data = Get-Content "06_VoiceRouting.json" | ConvertFrom-Json
$phoneNumbers = $data.PhoneNumberAssignments | ConvertTo-Csv -NoTypeInformation
$phoneNumbers | Out-File "PhoneNumbers.csv"
```

### Excel Analysis
```powershell
# Extract summary data to Excel-compatible format
$summary = Get-Content "00_Summary.json" | ConvertFrom-Json  
$summary.DataSummary | ConvertTo-Csv -NoTypeInformation | Out-File "DataSummary.csv"
```

### Automated Monitoring
```powershell
# Schedule monthly comprehensive analysis
$ScriptBlock = {
    Set-Location "C:\TeamsScripts"
    .\Gather-Teams-Calling-Data.ps1 -OutputPath "C:\Reports\$(Get-Date -Format 'yyyy-MM')" -GenerateCallFlowMaps -GeneratePDF -CompressOutput
}

$Trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At 6:00AM
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-Command & {$ScriptBlock}"  
Register-ScheduledTask -TaskName "Teams Monthly Analysis" -Trigger $Trigger -Action $Action
```

## 🎓 Advanced Usage Scenarios

### Scenario 1: Complete Configuration Audit
```powershell
# Comprehensive audit for compliance review
.\Gather-Teams-Calling-Data.ps1 `
  -OutputPath "C:\Audit\Teams_$(Get-Date -Format 'yyyy-MM-dd')" `
  -IncludeUserData `
  -GenerateCallFlowMaps `
  -GeneratePDF `
  -CompressOutput

# Result: Complete documentation package ready for auditors
```

### Scenario 2: Troubleshooting Specific Call Issues  
```powershell
# Focus on problem phone numbers for detailed analysis
.\Gather-Teams-Calling-Data.ps1 `
  -GenerateCallFlowMaps `
  -GeneratePDF `
  -CallFlowFilterNumbers "+15551234567", "+18005551234" 

# Result: Detailed call flow analysis for specific numbers
```

### Scenario 3: Migration Planning
```powershell
# Document current state before Teams calling migration
.\Gather-Teams-Calling-Data.ps1 `
  -OutputPath "C:\Migration\PreMigration_$(Get-Date -Format 'yyyy-MM-dd')" `
  -IncludeUserData `
  -GenerateCallFlowMaps `
  -GeneratePDF `
  -CompressOutput

# Result: Complete baseline documentation for migration planning
```

### Scenario 4: Multi-Tenant Analysis
```powershell
# Analyze multiple tenants (run separately for each)
$tenants = @("tenant1.onmicrosoft.com", "tenant2.onmicrosoft.com")
foreach ($tenant in $tenants) {
    .\Gather-Teams-Calling-Data.ps1 `
      -TenantId $tenant `
      -OutputPath "C:\MultiTenant\$tenant" `
      -GenerateCallFlowMaps `
      -GeneratePDF
}
```

## 📚 Additional Resources

### Microsoft Documentation
- [Teams Phone System Documentation](https://docs.microsoft.com/en-us/microsoftteams/cloud-voice-landing-page)
- [Teams PowerShell Module Reference](https://docs.microsoft.com/en-us/powershell/module/teams/)
- [Direct Routing Configuration](https://docs.microsoft.com/en-us/microsoftteams/direct-routing-landing-page)

### PowerShell Resources  
- [PowerShell 7 Installation](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [Teams Module Installation](https://docs.microsoft.com/en-us/microsoftteams/teams-powershell-install)

### Python Resources (for PDF Generation)
- [Python Installation](https://www.python.org/downloads/)
- [Playwright Documentation](https://playwright.dev/python/)

## 🆘 Support and Maintenance

### Version Information
- **Suite Version**: 2.0
- **PowerShell Compatibility**: 5.1+ and PowerShell 7+
- **Microsoft Teams Module**: 4.0.0+ required
- **Python Requirement**: 3.7+ (for PDF generation only)
- **Last Updated**: September 26, 2025
- **Maintenance**: Active development and bug fixes

### Getting Help
1. **Check troubleshooting section** above for common issues
2. **Review console output** for specific error messages  
3. **Enable verbose logging** for detailed diagnostics
4. **Verify prerequisites** are properly installed
5. **Check permissions** and authentication status

### Contributing
- **Issues and feedback**: Use repository issue tracker
- **Feature requests**: Submit detailed requirements
- **Code contributions**: Follow PowerShell best practices
- **Documentation**: Help improve this README

---

## 📄 License and Disclaimer

*This tool suite is designed for Microsoft Teams calling configuration analysis and troubleshooting. It is provided as-is for educational and professional use. Always follow your organization's security and compliance policies when collecting and analyzing Teams configuration data.*

## About

**Created for TMC Calling Issues analysis and Teams administration tasks.**

## System Requirements

- **PowerShell**: 5.1 or newer (PowerShell 7 recommended)
- **Microsoft Teams Module**: `Install-Module MicrosoftTeams`
- **Python**: 3.7 or higher (for PDF generation)
- **Playwright**: Modern browser automation (replaces discontinued wkhtmltopdf)
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

*This tool is designed for TMC Calling Issues analysis and Teams calling configuration auditing.*
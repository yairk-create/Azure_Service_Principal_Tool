# AzDevOpsVariableGroups PowerShell Module

A PowerShell module for managing Azure DevOps Variable Groups with Service Principal configurations. Streamline the creation and management of variable groups across multiple environments (dev, qa, prod) from CSV data.

## Features

- 🔧 **Automated Variable Group Creation** - Create variable groups from CSV data
- 🔐 **Secure Secret Management** - Properly handle Service Principal secrets
- 🌍 **Multi-Environment Support** - Deploy to dev, qa, prod environments
- ✅ **Validation & Testing** - Built-in validation and secret testing utilities
- 📊 **Comprehensive Reporting** - View and manage existing variable groups
- 🛡️ **Error Handling** - Robust error handling and logging

## Prerequisites

- **PowerShell 5.1** or later
- **Azure CLI** installed and configured
- **Azure DevOps CLI Extension** (auto-installed by the module)
- **Access to Azure DevOps** project with appropriate permissions

## Installation

### Option 1: Manual Installation

1. Download the module files:
   - `AzDevOpsVariableGroups.psm1`
   - `AzDevOpsVariableGroups.psd1`

2. Create the module directory:
   ```powershell
   $ModulePath = ($env:PSModulePath -split ';')[0]
   $ModuleDir = Join-Path $ModulePath "AzDevOpsVariableGroups"
   New-Item -Path $ModuleDir -ItemType Directory -Force
   ```

3. Copy both files to the module directory

4. Import the module:
   ```powershell
   Import-Module AzDevOpsVariableGroups -Force
   ```

### Option 2: Direct Installation Script

```powershell
# Run this script to automatically install the module
# (Replace with actual download/installation commands)
```

## CSV Data Format

Your CSV file must contain the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| `env` | Environment name | dev, qa, prod |
| `app` | Application name | MyApp-Service-DEV |
| `client_id` | Service Principal Client ID | f57b4f28-776b-422c-b39a-e28ed35815f0 |
| `tenant_id` | Azure Tenant ID | f4c80c7c-e1aa-4090-8a5d-c87dde95d0ee |
| `client_secret` | Service Principal Secret | M6T8Q~ueGRAiVxxeISH1tn_cdRxMpJycp.heUbN3 |

### Example CSV Content:
```csv
env,app,client_id,tenant_id,client_secret
dev,MyApp-API-DEV,f57b4f28-776b-422c-b39a-e28ed35815f0,f4c80c7c-e1aa-4090-8a5d-c87dde95d0ee,M6T8Q~ueGRAiVxxeISH1tn_cdRxMpJycp.heUbN3
dev,MyApp-Web-DEV,acefc269-d641-4db2-890e-bd2b103e7fe7,f4c80c7c-e1aa-4090-8a5d-c87dde95d0ee,N7U9R~vfHSBjWyyeJTI2uo_deSyNqKzdr.ifVcO4
qa,MyApp-API-QA,40be048f-f3cd-464c-97c4-f992cc661fff,f4c80c7c-e1aa-4090-8a5d-c87dde95d0ee,P8V0S~wgITCkXzzfKUJ3vp_gfTzOrL0es.jgWdP5
```

## Quick Start

### 1. Basic Usage

```powershell
# Import the module
Import-Module AzDevOpsVariableGroups

# Create variable groups for all environments
Set-EnvVariableGroups -Org "myorg" -Project "myproject" -Prefix "SPN" -CsvPath "C:\data\spns.csv"
```

### 2. Preview Changes (Recommended First)

```powershell
# See what would be created without making changes
Set-EnvVariableGroups -Org "myorg" -Project "myproject" -Prefix "SPN" -CsvPath "C:\data\spns.csv" -WhatIf
```

### 3. Single Environment

```powershell
# Process only dev environment
Set-EnvVariableGroups -Org "myorg" -Project "myproject" -Prefix "SPN" -CsvPath "C:\data\spns.csv" -Envs @('dev')
```

## Available Commands

### Set-EnvVariableGroups

Creates or updates Azure DevOps Variable Groups from CSV data.

**Parameters:**
- `-Org` (Required) - Azure DevOps Organization name
- `-Project` (Required) - Azure DevOps Project name  
- `-Prefix` (Required) - Prefix for variable group names
- `-CsvPath` (Required) - Path to CSV file with SPN details
- `-Envs` (Optional) - Environments to process (default: dev,qa,prod)
- `-DescTemplate` (Optional) - Description template for variable groups
- `-WhatIf` (Optional) - Preview changes without executing

**Examples:**
```powershell
# Basic usage
Set-EnvVariableGroups -Org "contoso" -Project "myapp" -Prefix "SPN" -CsvPath "spns.csv"

# Custom environments
Set-EnvVariableGroups -Org "contoso" -Project "myapp" -Prefix "API-SPN" -CsvPath "spns.csv" -Envs @('dev','staging')

# Preview mode
Set-EnvVariableGroups -Org "contoso" -Project "myapp" -Prefix "SPN" -CsvPath "spns.csv" -WhatIf
```

### Test-VariableGroupSecrets

Tests if secrets in a variable group are properly configured.

**Parameters:**
- `-GroupId` (Required) - Variable Group ID to test
- `-SecretNames` (Required) - Array of secret variable names to test

**Example:**
```powershell
Test-VariableGroupSecrets -GroupId "10" -SecretNames @("MyApp_ClientSecret", "AnotherApp_ClientSecret")
```

### Get-VariableGroupInfo

Lists variable groups in the current project.

**Parameters:**
- `-Filter` (Optional) - Filter variable groups by name

**Examples:**
```powershell
# List all variable groups
Get-VariableGroupInfo

# Filter by name
Get-VariableGroupInfo -Filter "SPN"
```

## Variable Group Structure

The module creates variable groups with the following naming convention:

**Variable Group Name:** `{Prefix}-{Environment}`
- Example: `SPN-dev`, `SPN-qa`, `SPN-prod`

**Variables Created per Application:**
- `{SafeAppName}_AppName` - Original application name
- `{SafeAppName}_ClientID` - Service Principal Client ID  
- `{SafeAppName}_TenantID` - Azure Tenant ID
- `{SafeAppName}_ClientSecret` - Service Principal Secret (marked as secret)

**Example Variables:**
```
MyApp-API-DEV_AppName      = "MyApp-API-DEV"
MyApp-API-DEV_ClientID     = "f57b4f28-776b-422c-b39a-e28ed35815f0"  
MyApp-API-DEV_TenantID     = "f4c80c7c-e1aa-4090-8a5d-c87dde95d0ee"
MyApp-API-DEV_ClientSecret = "***" (secret)
```

## Using Variables in Pipelines

Reference the variables in your Azure DevOps pipelines:

```yaml
# azure-pipelines.yml
variables:
- group: SPN-dev

steps:
- powershell: |
    Write-Host "Connecting to Azure..."
    az login --service-principal `
      --username $(MyApp-API-DEV_ClientID) `
      --password $(MyApp-API-DEV_ClientSecret) `
      --tenant $(MyApp-API-DEV_TenantID)
  displayName: 'Azure Login'
```

## Troubleshooting

### Common Issues

**1. Azure CLI Not Found**
```
Error: Azure CLI not found. Please install Azure CLI.
```
**Solution:** Install Azure CLI from https://docs.microsoft.com/en-us/cli/azure/install-azure-cli

**2. Missing CSV Columns**
```
Error: CSV missing required column: client_secret
```
**Solution:** Ensure your CSV has all required columns: env, app, client_id, tenant_id, client_secret

**3. Access Denied**
```
Error: You do not have permissions to create variable groups
```
**Solution:** Ensure you have "Variable Groups Administrator" permissions in Azure DevOps

**4. Organization/Project Not Found**
```
Error: Project 'myproject' was not found
```
**Solution:** Verify organization and project names are correct and you have access

### Debug Mode

Enable verbose output for troubleshooting:

```powershell
Set-EnvVariableGroups -Org "myorg" -Project "myproject" -Prefix "SPN" -CsvPath "spns.csv" -Verbose
```

### Test Secrets

Verify secrets are properly stored:

```powershell
# Test specific secrets
Test-VariableGroupSecrets -GroupId "10" -SecretNames @("MyApp_ClientSecret")
```

## Best Practices

### Security
- ✅ **Store CSV files securely** - Don't commit secrets to source control
- ✅ **Use strong passwords** - Rotate Service Principal secrets regularly  
- ✅ **Limit access** - Only grant necessary permissions to variable groups
- ✅ **Audit regularly** - Review variable group access and contents

### Organization
- ✅ **Consistent naming** - Use clear, consistent prefixes and naming conventions
- ✅ **Environment separation** - Keep dev/qa/prod variable groups separate
- ✅ **Documentation** - Document what each variable group contains
- ✅ **Version control** - Track changes to variable group configurations

### Automation
- ✅ **Use WhatIf** - Always preview changes before executing
- ✅ **Test first** - Test in dev environment before deploying to production
- ✅ **Automate updates** - Consider automating secret rotation
- ✅ **Monitor usage** - Track which pipelines use which variable groups

## Advanced Usage

### Custom Description Template

```powershell
$customTemplate = "Service Principals for {apps} applications in {env} environment - Updated $(Get-Date -Format 'yyyy-MM-dd')"

Set-EnvVariableGroups -Org "myorg" -Project "myproject" -Prefix "SPN" -CsvPath "spns.csv" -DescTemplate $customTemplate
```

### Bulk Processing

```powershell
# Process multiple CSV files
$csvFiles = Get-ChildItem "C:\data\*.csv"
foreach ($csv in $csvFiles) {
    $envName = $csv.BaseName
    Set-EnvVariableGroups -Org "myorg" -Project "myproject" -Prefix "SPN" -CsvPath $csv.FullName -Envs @($envName)
}
```

### Integration with CI/CD

```powershell
# Example pipeline script
param(
    [string]$Environment = "dev",
    [string]$CsvPath = "secrets.csv"
)

Import-Module AzDevOpsVariableGroups

# Validate first
Set-EnvVariableGroups -Org $env:AZURE_DEVOPS_ORG -Project $env:AZURE_DEVOPS_PROJECT -Prefix "SPN" -CsvPath $CsvPath -Envs @($Environment) -WhatIf

# Deploy if validation passes
if ($LASTEXITCODE -eq 0) {
    Set-EnvVariableGroups -Org $env:AZURE_DEVOPS_ORG -Project $env:AZURE_DEVOPS_PROJECT -Prefix "SPN" -CsvPath $CsvPath -Envs @($Environment)
}
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Update documentation
6. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

- **Documentation:** See inline help with `Get-Help <command> -Full`
- **Issues:** Report issues on the project repository
- **Questions:** Contact the development team

## Changelog

### Version 1.0.0 (2025-01-XX)
- Initial release
- Create/update Azure DevOps Variable Groups from CSV
- Support for multiple environments (dev, qa, prod)
- Secret management for Service Principals  
- Testing utilities for variable groups
- Comprehensive error handling and validation

---

**Made with ❤️ for Azure DevOps automation**
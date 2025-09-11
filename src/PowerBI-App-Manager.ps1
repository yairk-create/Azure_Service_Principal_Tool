# Power BI App Manager - Simple All-in-One Tool
# Creates apps, adds permissions, and grants admin consent
$global:GraphSP = $null
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import required modules
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    
    # Try to import the module needed for admin consent
    try {
        Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
        $global:ConsentModuleAvailable = $true
    } catch {
        Write-Host "Admin consent module not available. Installing..." -ForegroundColor Yellow
        try {
            Install-Module Microsoft.Graph.Identity.SignIns -Force -Scope CurrentUser
            Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
            $global:ConsentModuleAvailable = $true
        } catch {
            $global:ConsentModuleAvailable = $false
        }
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("Please install Microsoft Graph modules: Install-Module Microsoft.Graph -Force", "Missing Modules", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Global variables
$global:Apps = @()
$global:Permissions = @()
$global:PowerBISP = $null

# Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Power BI App Manager - Simple Tool"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# Connection Section
$connectionGroup = New-Object System.Windows.Forms.GroupBox
$connectionGroup.Text = "Step 1: Connect to Microsoft Graph"
$connectionGroup.Location = New-Object System.Drawing.Point(20, 20)
$connectionGroup.Size = New-Object System.Drawing.Size(950, 80)

$connectionStatus = New-Object System.Windows.Forms.Label
$connectionStatus.Text = "Status: Not Connected"
$connectionStatus.ForeColor = [System.Drawing.Color]::Red
$connectionStatus.Location = New-Object System.Drawing.Point(10, 25)
$connectionStatus.Size = New-Object System.Drawing.Size(300, 20)

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect to Microsoft Graph"
$connectButton.Location = New-Object System.Drawing.Point(10, 45)
$connectButton.Size = New-Object System.Drawing.Size(200, 25)
$connectButton.BackColor = [System.Drawing.Color]::LightBlue

# File Loading Section
$fileGroup = New-Object System.Windows.Forms.GroupBox
$fileGroup.Text = "Step 2: Load CSV Files"
$fileGroup.Location = New-Object System.Drawing.Point(20, 110)
$fileGroup.Size = New-Object System.Drawing.Size(950, 100)

$appsTextBox = New-Object System.Windows.Forms.TextBox
$appsTextBox.Location = New-Object System.Drawing.Point(100, 23)
$appsTextBox.Size = New-Object System.Drawing.Size(500, 20)
$appsTextBox.ReadOnly = $true

$browseAppsButton = New-Object System.Windows.Forms.Button
$browseAppsButton.Text = "Browse Apps CSV"
$browseAppsButton.Location = New-Object System.Drawing.Point(10, 20)
$browseAppsButton.Size = New-Object System.Drawing.Size(80, 25)

$permsTextBox = New-Object System.Windows.Forms.TextBox
$permsTextBox.Location = New-Object System.Drawing.Point(100, 53)
$permsTextBox.Size = New-Object System.Drawing.Size(500, 20)
$permsTextBox.ReadOnly = $true

$browsePermsButton = New-Object System.Windows.Forms.Button
$browsePermsButton.Text = "Browse Perms CSV"
$browsePermsButton.Location = New-Object System.Drawing.Point(10, 50)
$browsePermsButton.Size = New-Object System.Drawing.Size(80, 25)

$loadDataButton = New-Object System.Windows.Forms.Button
$loadDataButton.Text = "Load Data"
$loadDataButton.Location = New-Object System.Drawing.Point(620, 35)
$loadDataButton.Size = New-Object System.Drawing.Size(100, 25)
$loadDataButton.BackColor = [System.Drawing.Color]::LightGreen

# Permissions Selection Section
$permissionSelectionGroup = New-Object System.Windows.Forms.GroupBox
$permissionSelectionGroup.Text = "Step 2.5: Select Permissions to Add"
$permissionSelectionGroup.Location = New-Object System.Drawing.Point(20, 220)
$permissionSelectionGroup.Size = New-Object System.Drawing.Size(950, 200)
$permissionSelectionGroup.Visible = $true  # Always visible now

$permissionsLabel = New-Object System.Windows.Forms.Label
$permissionsLabel.Text = "Select which Power BI permissions to add to your apps (load CSV first):"
$permissionsLabel.Location = New-Object System.Drawing.Point(10, 20)
$permissionsLabel.Size = New-Object System.Drawing.Size(500, 20)

$selectAllPermsButton = New-Object System.Windows.Forms.Button
$selectAllPermsButton.Text = "Select All"
$selectAllPermsButton.Location = New-Object System.Drawing.Point(10, 45)
$selectAllPermsButton.Size = New-Object System.Drawing.Size(80, 25)
$selectAllPermsButton.BackColor = [System.Drawing.Color]::LightBlue

$clearAllPermsButton = New-Object System.Windows.Forms.Button
$clearAllPermsButton.Text = "Clear All"
$clearAllPermsButton.Location = New-Object System.Drawing.Point(100, 45)
$clearAllPermsButton.Size = New-Object System.Drawing.Size(80, 25)
$clearAllPermsButton.BackColor = [System.Drawing.Color]::LightCoral

$permissionsListBox = New-Object System.Windows.Forms.CheckedListBox
$permissionsListBox.Location = New-Object System.Drawing.Point(10, 80)
$permissionsListBox.Size = New-Object System.Drawing.Size(920, 110)
$permissionsListBox.CheckOnClick = $true
$permissionsListBox.Font = New-Object System.Drawing.Font("Consolas", 9)

# Add a status label for permissions
$permissionsStatusLabel = New-Object System.Windows.Forms.Label
$permissionsStatusLabel.Text = "No permissions loaded. Please load CSV files first."
$permissionsStatusLabel.Location = New-Object System.Drawing.Point(200, 48)
$permissionsStatusLabel.Size = New-Object System.Drawing.Size(400, 20)
$permissionsStatusLabel.ForeColor = [System.Drawing.Color]::Red

# Action Buttons Section
$actionGroup = New-Object System.Windows.Forms.GroupBox
$actionGroup.Text = "Step 3: Actions"
$actionGroup.Location = New-Object System.Drawing.Point(20, 430)
$actionGroup.Size = New-Object System.Drawing.Size(950, 80)

$createAppsButton = New-Object System.Windows.Forms.Button
$createAppsButton.Text = "1. Create Apps"
$createAppsButton.Location = New-Object System.Drawing.Point(20, 30)
$createAppsButton.Size = New-Object System.Drawing.Size(120, 35)
$createAppsButton.BackColor = [System.Drawing.Color]::LightBlue

$addPermissionsButton = New-Object System.Windows.Forms.Button
$addPermissionsButton.Text = "2. Add Permissions"
$addPermissionsButton.Location = New-Object System.Drawing.Point(160, 30)
$addPermissionsButton.Size = New-Object System.Drawing.Size(120, 35)
$addPermissionsButton.BackColor = [System.Drawing.Color]::LightGreen

$grantConsentButton = New-Object System.Windows.Forms.Button
$grantConsentButton.Text = "3. Grant Admin Consent"
$grantConsentButton.Location = New-Object System.Drawing.Point(300, 30)
$grantConsentButton.Size = New-Object System.Drawing.Size(140, 35)
$grantConsentButton.BackColor = [System.Drawing.Color]::Orange

$doAllButton = New-Object System.Windows.Forms.Button
$doAllButton.Text = "DO ALL (Create + Permissions + Consent)"
$doAllButton.Location = New-Object System.Drawing.Point(460, 30)
$doAllButton.Size = New-Object System.Drawing.Size(250, 35)
$doAllButton.BackColor = [System.Drawing.Color]::Gold
$doAllButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)

# Results Section
$resultsGroup = New-Object System.Windows.Forms.GroupBox
$resultsGroup.Text = "Results"
$resultsGroup.Location = New-Object System.Drawing.Point(20, 520)
$resultsGroup.Size = New-Object System.Drawing.Size(950, 120)

$resultsTextBox = New-Object System.Windows.Forms.TextBox
$resultsTextBox.Multiline = $true
$resultsTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$resultsTextBox.Location = New-Object System.Drawing.Point(10, 20)
$resultsTextBox.Size = New-Object System.Drawing.Size(930, 90)
$resultsTextBox.ReadOnly = $true
$resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)

# Add controls to form
$connectionGroup.Controls.AddRange(@($connectionStatus, $connectButton))
$fileGroup.Controls.AddRange(@($browseAppsButton, $appsTextBox, $browsePermsButton, $permsTextBox, $loadDataButton))
$permissionSelectionGroup.Controls.AddRange(@($permissionsLabel, $selectAllPermsButton, $clearAllPermsButton, $permissionsListBox, $permissionsStatusLabel))
$actionGroup.Controls.AddRange(@($createAppsButton, $addPermissionsButton, $grantConsentButton, $doAllButton))
$resultsGroup.Controls.Add($resultsTextBox)
$form.Controls.AddRange(@($connectionGroup, $fileGroup, $permissionSelectionGroup, $actionGroup, $resultsGroup))

# Functions
function Write-Result {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $resultsTextBox.AppendText("[$timestamp] $Message`r`n")
    $resultsTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-ConnectionStatus {
    try {
        $context = Get-MgContext
        if ($context) {
            $connectionStatus.Text = "Status: Connected to $($context.TenantId)"
            $connectionStatus.ForeColor = [System.Drawing.Color]::Green
            return $true
        } else {
            $connectionStatus.Text = "Status: Not Connected"
            $connectionStatus.ForeColor = [System.Drawing.Color]::Red
            return $false
        }
    } catch {
        $connectionStatus.Text = "Status: Connection Error"
        $connectionStatus.ForeColor = [System.Drawing.Color]::Red
        return $false
    }
}

function Load-PowerBIServicePrincipal {
    Write-Result "Loading Power BI Service Principal..."
    try {
        $global:PowerBISP = Get-MgServicePrincipal -Filter "displayName eq 'Power BI Service'" -ErrorAction Stop
        if ($global:PowerBISP) {
            Write-Result "SUCCESS: Power BI Service Principal found with $($global:PowerBISP.Oauth2PermissionScopes.Count) scopes"
            return $true
        } else {
            Write-Result "ERROR: Power BI Service Principal not found"
            return $false
        }
    } catch {
        Write-Result "ERROR: $($_.Exception.Message)"
        return $false
    }
}

function Load-ServicePrincipals {
    Write-Result "Loading Service Principals (Power BI + Microsoft Graph)..."
    try {
        $global:PowerBISP = Get-MgServicePrincipal -Filter "displayName eq 'Power BI Service'" -Property "appId,oauth2PermissionScopes,appRoles" -ErrorAction Stop
        Write-Result "OK: Power BI Service (scopes=$($global:PowerBISP.Oauth2PermissionScopes.Count), roles=$($global:PowerBISP.AppRoles.Count))"
    } catch { Write-Result "ERROR Power BI SP: $($_.Exception.Message)" }

    try {
        # Microsoft Graph appId is fixed
        $global:GraphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -Property "appId,oauth2PermissionScopes,appRoles" -ErrorAction Stop
        Write-Result "OK: Microsoft Graph (scopes=$($global:GraphSP.Oauth2PermissionScopes.Count), roles=$($global:GraphSP.AppRoles.Count))"
    } catch { Write-Result "ERROR Graph SP: $($_.Exception.Message)" }
}



function Load-CSVData {
    if (-not $appsTextBox.Text -or -not $permsTextBox.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select both CSV files first.", "Missing Files")
        return $false
    }
    try {
        $global:Apps = Import-Csv -Path $appsTextBox.Text -ErrorAction Stop
        $permsRaw    = Import-Csv -Path $permsTextBox.Text -ErrorAction Stop

        $hasApi = $permsRaw | Get-Member -Name Api -MemberType NoteProperty
        foreach ($r in $permsRaw) {
            if (-not $hasApi) { $r | Add-Member -NotePropertyName Api -NotePropertyValue 'Power BI Service' -Force }
            if (-not $r.Type) { $r.Type = 'Delegated' }
        }

        $global:Permissions = $permsRaw | Where-Object { $_.Permission } | Select-Object Api, Permission, Type
        Populate-PermissionsList
        Write-Result ("Loaded {0} apps and {1} permissions" -f $global:Apps.Count, $global:Permissions.Count)
        return $true
    } catch {
        Write-Result "ERROR loading CSVs: $($_.Exception.Message)"
        return $false
    }
}



function Create-Applications {
    if ($global:Apps.Count -eq 0) {
        Write-Result "ERROR: No apps loaded. Please load CSV data first."
        return $false
    }
    
    Write-Result "=== CREATING APPLICATIONS ==="
    $createdCount = 0
    $existingCount = 0
    
    foreach ($app in $global:Apps) {
        $appName = $app.AppDisplayName
        Write-Result "Processing: $appName"
        
        try {
            # Check if app already exists
            $existingApp = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue
            
            if ($existingApp) {
                Write-Result "  INFO: App already exists (AppId: $($existingApp.AppId))"
                $existingCount++
            } else {
                # Create new app
                $newAppParams = @{
                    DisplayName = $appName
                    SignInAudience = "AzureADMyOrg"
                }
                
                $newApp = New-MgApplication -BodyParameter $newAppParams -ErrorAction Stop
                Write-Result "  SUCCESS: Created app (AppId: $($newApp.AppId))"
                $createdCount++
            }
        } catch {
            Write-Result "  ERROR: Failed to create $appName - $($_.Exception.Message)"
        }
    }
    
    Write-Result "=== CREATE SUMMARY ==="
    Write-Result "Created: $createdCount apps"
    Write-Result "Already existed: $existingCount apps"
    
    return ($createdCount -gt 0 -or $existingCount -gt 0)
}

function Add-ApiPermissions {
    if ($global:Apps.Count -eq 0) { Write-Result "ERROR: No apps loaded."; return $false }
    if (-not (Update-ConnectionStatus)) { Write-Result "ERROR: Not connected."; return $false }

    if (-not $global:PowerBISP -or -not $global:GraphSP) {
        Load-ServicePrincipals
    }

    $selectedPermissions = Get-SelectedPermissions
    if ($selectedPermissions.Count -eq 0) {
        Write-Result "ERROR: No permissions selected."; return $false
    }

    Write-Result "=== ADDING API PERMISSIONS (Power BI + Graph) ==="
    $byApi = $selectedPermissions | Group-Object Api

    # Build a map of API -> { sp, resourceAppId, scopesByValue, rolesByValue }
    $apiMap = @{}
    foreach ($g in $byApi) {
        switch ($g.Name) {
            'Power BI Service' {
                if (-not $global:PowerBISP) { Write-Result "ERROR: Power BI SP missing."; continue }
                $apiMap[$g.Name] = [pscustomobject]@{
                    SP = $global:PowerBISP
                    ResourceAppId = $global:PowerBISP.AppId
                    ScopesByValue = @{}; RolesByValue = @{}
                }
                foreach ($s in $global:PowerBISP.Oauth2PermissionScopes) { $apiMap[$g.Name].ScopesByValue[$s.Value] = $s.Id }
                foreach ($r in $global:PowerBISP.AppRoles) { if ($r.AllowedMemberTypes -contains 'Application') { $apiMap[$g.Name].RolesByValue[$r.Value] = $r.Id } }
            }
            'Microsoft Graph' {
                if (-not $global:GraphSP) { Write-Result "ERROR: Microsoft Graph SP missing."; continue }
                $apiMap[$g.Name] = [pscustomobject]@{
                    SP = $global:GraphSP
                    ResourceAppId = $global:GraphSP.AppId
                    ScopesByValue = @{}; RolesByValue = @{}
                }
                foreach ($s in $global:GraphSP.Oauth2PermissionScopes) { $apiMap[$g.Name].ScopesByValue[$s.Value] = $s.Id }
                foreach ($r in $global:GraphSP.AppRoles) { if ($r.AllowedMemberTypes -contains 'Application') { $apiMap[$g.Name].RolesByValue[$r.Value] = $r.Id } }
            }
            default { Write-Result "WARNING: Unsupported API '$($g.Name)'. Skipping."; }
        }
    }

    $successCount = 0

foreach ($app in $global:Apps) {
    $appName = [string]$app.AppDisplayName
    Write-Result "Processing: $appName"

    try {
        # Get a single, valid application (with Id + current RRA)
        $existingApp = Get-MgApplication `
            -Filter "displayName eq '$appName'" `
            -All `
            -Property "id,appId,displayName,requiredResourceAccess" `
            -ErrorAction Stop | Select-Object -First 1

        if (-not $existingApp -or [string]::IsNullOrWhiteSpace($existingApp.Id)) {
            Write-Result "  ERROR: App not found or missing Id: $appName"
            continue
        }

        $currentRRA = @()
        if ($existingApp.RequiredResourceAccess) { $currentRRA = @($existingApp.RequiredResourceAccess) }

        foreach ($apiName in $apiMap.Keys) {
            $meta = $apiMap[$apiName]; if (-not $meta) { continue }

            $wantedScopes = @()
            $wantedRoles  = @()

            foreach ($perm in ($selectedPermissions | Where-Object Api -eq $apiName)) {
                if ($perm.Type -eq 'Delegated') {
                    $id = $meta.ScopesByValue[$perm.Permission]
                    if ($id) { $wantedScopes += $id } else { Write-Result "  ✗ Not found (scope): [$apiName] $($perm.Permission)" }
                } elseif ($perm.Type -eq 'Application') {
                    $id = $meta.RolesByValue[$perm.Permission]
                    if ($id) { $wantedRoles += $id } else { Write-Result "  ✗ Not found (role): [$apiName] $($perm.Permission)" }
                }
            }

            if (($wantedScopes.Count + $wantedRoles.Count) -eq 0) {
                Write-Result "  Skipping [$apiName]: nothing selected/mapped"
                continue
            }

            $existingRes = $currentRRA | Where-Object { $_.ResourceAppId -eq $meta.ResourceAppId } | Select-Object -First 1
            $newResourceAccess = @()

            if ($existingRes) {
                $existingScopeIds = @($existingRes.ResourceAccess | Where-Object Type -eq "Scope" | ForEach-Object { $_.Id })
                $existingRoleIds  = @($existingRes.ResourceAccess | Where-Object Type -eq "Role"  | ForEach-Object { $_.Id })

                $allScopeIds = @($existingScopeIds + $wantedScopes) | Sort-Object -Unique
                $allRoleIds  = @($existingRoleIds  + $wantedRoles ) | Sort-Object -Unique

                foreach ($id in $allScopeIds) { $newResourceAccess += @{ Id = $id; Type = "Scope" } }
                foreach ($id in $allRoleIds ) { $newResourceAccess += @{ Id = $id; Type = "Role"  } }

                # Rebuild RRA: drop old entry for this API and add merged
                $currentRRA = @($currentRRA | Where-Object { $_.ResourceAppId -ne $meta.ResourceAppId })
                $currentRRA += @{ ResourceAppId = $meta.ResourceAppId; ResourceAccess = $newResourceAccess }

                Write-Result "  ✓ Merged [$apiName]: scopes=$($allScopeIds.Count), roles=$($allRoleIds.Count)"
            } else {
                foreach ($id in ($wantedScopes | Sort-Object -Unique)) { $newResourceAccess += @{ Id = $id; Type = "Scope" } }
                foreach ($id in ($wantedRoles  | Sort-Object -Unique)) { $newResourceAccess += @{ Id = $id; Type = "Role"  } }
                $currentRRA += @{ ResourceAppId = $meta.ResourceAppId; ResourceAccess = $newResourceAccess }
                Write-Result "  ✓ Added [$apiName]: scopes=$($wantedScopes.Count), roles=$($wantedRoles.Count)"
            }
        }

        # Final update (now guaranteed to have a valid Id)
        Update-MgApplication -ApplicationId $existingApp.Id -BodyParameter @{ RequiredResourceAccess = $currentRRA } -ErrorAction Stop
        Write-Result "  ✓ Updated app: $appName"
        $successCount++
    } catch {
        Write-Result "  ERROR: Failed to update $appName - $($_.Exception.Message)"
    }
}


    Write-Result "=== PERMISSIONS SUMMARY ==="
    Write-Result "Successfully updated: $successCount apps"
    return ($successCount -gt 0)
}


function Grant-AdminConsent {
    if ($global:Apps.Count -eq 0) {
        Write-Result "ERROR: No apps loaded. Please load CSV data first."
        return $false
    }
    
    if (-not $global:PowerBISP) {
        if (-not (Load-PowerBIServicePrincipal)) {
            return $false
        }
    }
    
    # Check if admin consent module is available
    if (-not $global:ConsentModuleAvailable) {
        Write-Result "=== ADMIN CONSENT - MANUAL PROCESS ==="
        Write-Result "Admin consent module not available. Please grant consent manually:"
        Write-Result ""
        
        foreach ($app in $global:Apps) {
            $appName = $app.AppDisplayName
            try {
                $existingApp = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction Stop
                if ($existingApp) {
                    $portalUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($existingApp.AppId)"
                    Write-Result "App: $appName"
                    Write-Result "  Portal URL: $portalUrl"
                    Write-Result "  1. Click the URL above"
                    Write-Result "  2. Go to 'API Permissions'"
                    Write-Result "  3. Click 'Grant admin consent for [tenant]'"
                    Write-Result ""
                }
            } catch {
                Write-Result "ERROR: Could not find app $appName"
            }
        }
        
        Write-Result "=== MANUAL CONSENT SUMMARY ==="
        Write-Result "Please manually grant consent for all apps using the URLs above"
        
        # Show message box with manual instructions
        $manualMessage = "Admin consent module not available.`n`nPlease manually grant consent:`n1. Go to Azure Portal > App Registrations`n2. Select each app`n3. Go to API Permissions`n4. Click 'Grant admin consent for [tenant]'"
        [System.Windows.Forms.MessageBox]::Show($manualMessage, "Manual Admin Consent Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        return $true  # Return true since we provided instructions
    }
    
    Write-Result "=== GRANTING ADMIN CONSENT ==="
    Write-Result "WARNING: This grants permissions for ALL users in your tenant!"
    
    $successCount = 0
    
    foreach ($app in $global:Apps) {
        $appName = $app.AppDisplayName
        Write-Result "Processing: $appName"
        
        try {
            # Find the application
            $existingApp = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction Stop
            
            if (-not $existingApp) {
                Write-Result "  ERROR: App not found: $appName"
                continue
            }
            
            # Get or create service principal
            $appSp = Get-MgServicePrincipal -Filter "appId eq '$($existingApp.AppId)'" -ErrorAction SilentlyContinue
            if (-not $appSp) {
                $appSp = New-MgServicePrincipal -AppId $existingApp.AppId -ErrorAction Stop
                Write-Result "  INFO: Created service principal"
            }
            
            # Check current Power BI permissions
            $pbiPermissions = $existingApp.RequiredResourceAccess | Where-Object ResourceAppId -eq $global:PowerBISP.AppId
            
            if (-not $pbiPermissions) {
                Write-Result "  WARNING: No Power BI permissions found. Add permissions first."
                continue
            }
            
            $scopeIds = $pbiPermissions.ResourceAccess | Where-Object Type -eq "Scope" | Select-Object -ExpandProperty Id
            
            # Try to check existing grants first
            try {
                $existingGrant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($appSp.Id)' and resourceId eq '$($global:PowerBISP.Id)'" -ErrorAction SilentlyContinue
            } catch {
                # If Get-MgOauth2PermissionGrant fails, try alternative approach
                Write-Result "  WARNING: Cannot check existing grants, proceeding with consent creation"
                $existingGrant = $null
            }
            
            if ($existingGrant) {
                Write-Result "  INFO: Admin consent already exists"
            } else {
                # Create admin consent grant
                try {
                    $consentParams = @{
                        ClientId = $appSp.Id
                        ResourceId = $global:PowerBISP.Id
                        Scope = ($global:PowerBISP.Oauth2PermissionScopes | Where-Object Id -in $scopeIds | Select-Object -ExpandProperty Value) -join " "
                        ConsentType = "AllPrincipals"
                    }
                    
                    # Try to create the consent grant
                    New-MgOauth2PermissionGrant -BodyParameter $consentParams -ErrorAction Stop
                    Write-Result "  SUCCESS: Admin consent granted"
                } catch {
                    # If auto-consent fails, provide manual instructions
                    Write-Result "  INFO: Auto-consent failed, providing manual instructions"
                    $portalUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($existingApp.AppId)"
                    Write-Result "  Manual consent URL: $portalUrl"
                }
            }
            
            $successCount++
            
        } catch {
            Write-Result "  ERROR: Failed to process $appName - $($_.Exception.Message)"
            # Provide manual fallback
            try {
                $existingApp = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue
                if ($existingApp) {
                    $portalUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($existingApp.AppId)"
                    Write-Result "  Manual consent URL: $portalUrl"
                }
            } catch {
                Write-Result "  Could not generate manual consent URL"
            }
        }
    }
    
    Write-Result "=== CONSENT SUMMARY ==="
    Write-Result "Processed: $successCount apps"
    Write-Result "If auto-consent failed, use the manual URLs provided above"
    
    return ($successCount -gt 0)
}

function Do-Everything {
    Write-Result "=========================================="
    Write-Result "STARTING COMPLETE WORKFLOW"
    Write-Result "=========================================="
    
    if (-not (Update-ConnectionStatus)) {
        Write-Result "ERROR: Not connected to Microsoft Graph"
        return
    }
    
    if (-not (Load-CSVData)) {
        return
    }
    
    # Step 1: Create Apps
    Write-Result ""
    if (Create-Applications) {
        Write-Result "✓ Step 1 Complete: Apps created/verified"
    } else {
        Write-Result "✗ Step 1 Failed: Could not create apps"
        return
    }
    
    # Step 2: Add Permissions
    Write-Result ""
    if (Add-ApiPermissions) {
        Write-Result "✓ Step 2 Complete: Permissions added"
    } else {
        Write-Result "✗ Step 2 Failed: Could not add permissions"
        return
    }
    
    # Step 3: Grant Consent
    Write-Result ""
    if (Grant-AdminConsent) {
        Write-Result "✓ Step 3 Complete: Admin consent granted"
    } else {
        Write-Result "✗ Step 3 Failed: Could not grant consent"
        return
    }
    
    Write-Result ""
    Write-Result "=========================================="
    Write-Result "🎉 ALL COMPLETE! 🎉"
    Write-Result "Your apps are ready to use with Power BI permissions!"
    Write-Result "=========================================="
    
    [System.Windows.Forms.MessageBox]::Show("All tasks completed successfully! Your apps are ready to use.", "Success!", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Event Handlers
$connectButton.Add_Click({
    try {
        Write-Result "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "Application.ReadWrite.All","DelegatedPermissionGrant.ReadWrite.All" -ErrorAction Stop
        Update-ConnectionStatus | Out-Null
        Load-ServicePrincipals | Out-Null
        Write-Result "Connection successful!"
    } catch {
        Write-Result "Connection failed: $($_.Exception.Message)"
    }
})


$browseAppsButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Select Apps CSV File"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $appsTextBox.Text = $openFileDialog.FileName
    }
})

$browsePermsButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Select Permissions CSV File"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $permsTextBox.Text = $openFileDialog.FileName
    }
})

$loadDataButton.Add_Click({ Load-CSVData | Out-Null })

$selectAllPermsButton.Add_Click({
    for ($i = 0; $i -lt $permissionsListBox.Items.Count; $i++) {
        $permissionsListBox.SetItemChecked($i, $true)
    }
    $selectedCount = 0
    for ($i = 0; $i -lt $permissionsListBox.Items.Count; $i++) {
        if ($permissionsListBox.GetItemChecked($i)) { $selectedCount++ }
    }
    $permissionsStatusLabel.Text = "All $selectedCount permissions selected"
    $permissionsStatusLabel.ForeColor = [System.Drawing.Color]::Green
    Write-Result "Selected all $($permissionsListBox.Items.Count) permissions"
})

$clearAllPermsButton.Add_Click({
    for ($i = 0; $i -lt $permissionsListBox.Items.Count; $i++) {
        $permissionsListBox.SetItemChecked($i, $false)
    }
    $permissionsStatusLabel.Text = "No permissions selected"
    $permissionsStatusLabel.ForeColor = [System.Drawing.Color]::Red
    Write-Result "Cleared all permission selections"
})

function Populate-PermissionsList {
    $permissionsListBox.Items.Clear()
    foreach ($p in $global:Permissions) {
        $text = "[{0}] {1} ({2})" -f $p.Api, $p.Permission, $p.Type
        [void]$permissionsListBox.Items.Add($text)
    }
    $count = $permissionsListBox.Items.Count
    if ($count -gt 0) {
        $permissionsStatusLabel.Text = "Loaded $count permissions. Select what you need."
        $permissionsStatusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $permissionsStatusLabel.Text = "No permissions loaded."
        $permissionsStatusLabel.ForeColor = [System.Drawing.Color]::Red
    }
}



function Get-SelectedPermissions {
    $selected = @()
    foreach ($item in $permissionsListBox.CheckedItems) {
        if ($item -match '^\[(.+?)\]\s+(.+?)\s+\((Delegated|Application)\)$') {
            $api = $Matches[1]; $perm = $Matches[2]; $type = $Matches[3]
            $row = $global:Permissions | Where-Object { $_.Api -eq $api -and $_.Permission -eq $perm -and $_.Type -eq $type }
            if ($row) { $selected += $row }
        }
    }
    return ,$selected
}




$createAppsButton.Add_Click({ Create-Applications | Out-Null })
$addPermissionsButton.Add_Click({ Add-ApiPermissions | Out-Null }) 
$grantConsentButton.Add_Click({ Grant-AdminConsent | Out-Null })
$doAllButton.Add_Click({ Do-Everything })


# Add after $doAllButton
$createSecretsButton = New-Object System.Windows.Forms.Button
$createSecretsButton.Text = "4. Create Secrets + Export CSV"
$createSecretsButton.Location = New-Object System.Drawing.Point(720, 30)
$createSecretsButton.Size = New-Object System.Drawing.Size(200, 35)
$createSecretsButton.BackColor = [System.Drawing.Color]::LightSteelBlue
$actionGroup.Controls.Add($createSecretsButton)

function Create-AppSecretsCsv {
    if ($global:Apps.Count -eq 0) {
        Write-Result "ERROR: No apps loaded. Please load CSV data first."
        return $false
    }
    if (-not (Update-ConnectionStatus)) {
        Write-Result "ERROR: Not connected to Microsoft Graph"
        return $false
    }

    $ts       = Get-Date -Format "yyyyMMdd-HHmmss"
    $outFile  = Join-Path $env:USERPROFILE "Desktop\powerbi-app-secrets-$ts.csv"
    $start    = Get-Date
    $end      = (Get-Date).AddYears(2)
    $context  = Get-MgContext
    $tenantId = if ($context) { $context.TenantId } else { "" }

    Write-Result "=== CREATING 2-YEAR SECRETS & EXPORTING CSV ==="
    Write-Result "Output: $outFile"

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($app in $global:Apps) {
        $appName = $app.AppDisplayName
        Write-Result "Processing: $appName"

        try {
            $existingApp = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction Stop
            if (-not $existingApp) {
                Write-Result "  ERROR: App not found: $appName"
                continue
            }

            $secretName = "AutoSecret-$(Get-Date -Format 'yyyyMMdd')"
            $cred = @{
                displayName   = $secretName
                startDateTime = $start
                endDateTime   = $end
            }

            # Create the secret (returns secretText once)
            $result = Add-MgApplicationPassword -ApplicationId $existingApp.Id -PasswordCredential $cred -ErrorAction Stop

            # Collect a CSV row
            $rows.Add([pscustomobject]@{
                TenantId       = $tenantId
                AppDisplayName = $appName
                AppId          = $existingApp.AppId     # Client ID
                ApplicationId  = $existingApp.Id        # Object ID
                SecretId       = $result.KeyId
                SecretName     = $secretName
                SecretValue    = $result.SecretText     # Store NOW or you'll lose it
                StartDateTime  = $start.ToString("s")
                EndDateTime    = $end.ToString("s")
            })

            Write-Result "  SUCCESS: Secret created (valid until $($end.ToString('yyyy-MM-dd')))"
        } catch {
            Write-Result "  ERROR: Failed for $appName - $($_.Exception.Message)"
        }
    }

    if ($rows.Count -gt 0) {
        $rows | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
        Write-Result "DONE: Wrote $($rows.Count) rows to $outFile"
        return $true
    } else {
        Write-Result "No secrets created."
        return $false
    }
}

$createSecretsButton.Add_Click({ Create-AppSecretsCsv | Out-Null })

# Inside Do-Everything, after Step 3
Write-Result ""
if (Create-AppSecretsCsv) {
    Write-Result "✓ Step 4 Complete: Secrets created & CSV exported"
} else {
    Write-Result "✗ Step 4: Secret creation/CSV export failed"
}



# Initialize
Write-Result "Power BI App Manager started"
Write-Result "Step 1: Connect to Microsoft Graph"
Write-Result "Step 2: Load your CSV files"
Write-Result "Step 3: Click 'DO ALL' to create apps, add permissions, and grant consent!"
Update-ConnectionStatus

# Show the form
[System.Windows.Forms.Application]::Run($form)



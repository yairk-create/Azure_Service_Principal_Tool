# Power BI App Manager

A simple all-in-one tool to automate creating Azure AD app registrations for Power BI, adding Microsoft Graph and Power BI API permissions, granting admin consent, and exporting client secrets to CSV.

---

## Features
- Create multiple App Registrations from a CSV list
- Add delegated and application API permissions (Power BI + Microsoft Graph)
- Grant admin consent (automatic or manual fallback with portal links)
- Generate 2-year client secrets and export them into a CSV
- GUI-based workflow (WinForms)

---

## Prerequisites
- Windows + PowerShell 5.1 or PowerShell 7.x
- .NET WinForms (included with Windows)
- Microsoft Graph modules:

  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
  ```

- Azure AD permissions:
  - Ability to create App Registrations
  - Tenant admin rights for admin consent

---

## CSV Templates

### `apps.csv`
```csv
AppDisplayName
My-PowerBI-App-001
My-PowerBI-App-002
```

### `perms.csv`
```csv
Api,Permission,Type
Power BI Service,Dataset.Read.All,Application
Power BI Service,Dataset.ReadWrite.All,Application
Microsoft Graph,User.Read,Delegated
```

---

## Usage
1. Clone this repository.
2. Run the script:

   ```powershell
   .\src\PowerBI-App-Manager.ps1
   ```

3. Follow the steps in the GUI:
   - Connect to Microsoft Graph
   - Load your CSVs
   - Create apps
   - Add permissions
   - Grant admin consent
   - Create secrets & export CSV

---

## Output
- A secrets file saved to your Desktop:

  ```
  powerbi-app-secrets-YYYYMMDD-HHMMSS.csv
  ```

This file contains:
- Tenant ID
- App display name
- AppId (Client ID)
- Object ID
- Secret ID
- Secret value
- Validity dates

---

## Notes
- If automatic admin consent fails, portal links will be generated for manual consent.
- Re-running is idempotent: apps already created or permissions already assigned are not duplicated.

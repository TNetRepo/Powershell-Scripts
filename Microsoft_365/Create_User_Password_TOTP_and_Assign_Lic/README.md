A PowerShell script that automates the provisioning of Microsoft 365 users, passwords, and assigning a Microsoft software license. The script creates user accounts, assigns defined Microsoft license, and configures Hardware OATH TOTP multi-factor authentication — all in a single run.

---

## What the Script Does

For each user defined in the `$users` array, the script performs the following steps in sequence:

1. **Creates the M365 user account** with the specified display name, first name, last name, UPN, and a randomly generated 16-character password. The password meets complexity requirements (uppercase, lowercase, number, special character) and does not require a reset on first sign-in, as these accounts are admin-managed.

2. **Assigns a Microsoft license** to the newly created user via the Microsoft Graph `assignLicense` endpoint.

3. **Generates a unique base32-encoded TOTP secret** and registers it as a Hardware OATH token device against the user account using the Microsoft Graph `hardwareOathDevices` beta endpoint.

4. **Retrieves the method ID** of the registered token from the user's assigned hardware OATH methods.

5. **Activates the TOTP token** by computing a live TOTP code from the generated secret and submitting it to the activation endpoint. The script includes logic to avoid TOTP window boundary timing issues.

On completion, a CSV file is written to the desktop containing each user's UPN, password, TOTP secret, and provisioning status. Any users that fail during processing are captured in the CSV with their error details — the script continues processing remaining users if an individual user fails.

---

## Prerequisites

### PowerShell Module

Install the Microsoft Graph PowerShell module before running the script:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Alternatively, install only the required sub-modules:

```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

### Entra Admin Center Configuration

Before running the script, ensure the following authentication method policies are enabled in the Entra admin center under **Protection → Authentication methods**:

- **Software OATH tokens** — must be enabled for the token registration step to succeed
- **Hardware OATH tokens (Preview)** — must be enabled for the registered tokens to be usable at sign-in

### Admin Role Requirements

The account used to connect must have one of the following roles in the target tenant:

- Global Administrator, or
- User Administrator + Authentication Policy Administrator

---

## Connecting to the Tenant

Run the following command before executing the script. Replace `your-tenant-id` with the tenant ID found in the Entra admin center under **Identity → Overview**.

```powershell
Connect-MgGraph -TenantId "your-tenant-id" -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All", "Organization.Read.All", "Policy.ReadWrite.AuthenticationMethod"
```

The connection persists for the duration of the PowerShell session. To verify you are connected to the correct tenant before running:

```powershell
Get-MgContext
```

---

## Variables to Configure Before Running

| Variable | Location | Description |
|---|---|---|
| `$domain` | Line 9 | The M365 domain suffix for user UPNs. Use your verified custom domain or the `.onmicrosoft.com` domain. Example: `contoso.onmicrosoft.com` |
| `$usageLocation` | Line 10 | Two-letter country code for license assignment. Example: `US` |
| `$pbiSkuId` | Line 11 | The GUID of the Microsoft license in your tenant. Retrieve this by running `Get-MgSubscribedSku \| Select SkuPartNumber, SkuId` and locating the row with the product SKU you want to assign. |
| `$users` | Line 88 | Array of user objects to provision. Each entry requires `Display`, `First`, `Last`, and `User` (the username prefix, without the domain). |

### Finding Your Microsoft License SKU ID (run this and find the SKU for the product you want to assign in the list)

```powershell
Get-MgSubscribedSku | Select SkuPartNumber, SkuId
```

### User Array Format

```powershell
$users = @(
    @{ Display="Jane Smith"; First="Jane"; Last="Smith"; User="JSmith" },
    @{ Display="John Doe";   First="John"; Last="Doe";   User="JDoe"   }
)
```

Username convention used in this deployment: first initial + last name (e.g. Bill Smith → `BSmith`).

---

## Output

On completion the script writes a CSV file to the current user's desktop:

```
C:\Users\<username>\Desktop\PBI_Users_Credentials.csv
```

The CSV contains the following columns:

| Column | Description |
|---|---|
| UPN | The user's full M365 login address |
| Password | The auto-generated 16-character password |
| TOTPSecret | The base32 TOTP secret for import into an authenticator app or password manager |
| Status | `Success` or `FAILED: <error detail>` |

> **Security note:** The CSV contains plaintext passwords and TOTP secrets. Import credentials into your password manager (e.g. Hudu) immediately after the script completes and delete the CSV file from the desktop.

---

## Execution Policy

If PowerShell blocks the script due to execution policy restrictions, run the following before executing — this relaxes the policy for the current session only:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## Re-running for Failed Users

If any users show `FAILED` in the output CSV, isolate them by temporarily replacing the `$users` array with only the failed entries and re-run the script. Verify in the Entra admin center that no partial account was created for that user before re-running to avoid duplicate UPN errors.

# Migration Manager for Exchange Online – Permissions & Usage Guide

## Overview

**Migration Manager** is a tool designed to automate user onboarding for Exchange Online and Microsoft 365.  
Key functions include:
- Sending Temporary Access Passes (TAP) to users for secure initial authentication and setup.
- Enabling migration of users into Exchange Online.
- Managing permissions for app-only authentication with Exchange Online via an Azure AD service principal (Enterprise Application).

---

## Features

- **Automated TAP Delivery:**  
  Sends personalized emails to users with their TAP, onboarding instructions, validity, and support contacts.

- **Streamlined User Migration:**  
  Once TAP is used, Migration Manager can initiate migration of user mailboxes and settings.

- **App-Only Exchange Online Admin:**  
  Uses an Azure AD app registration (service principal) assigned to the "Organization Management" role group for secure, scalable migration operations.

---

## Permissions Setup (Service Principal)

**Important:**  
For Exchange Online permissions, always use the Object ID from the Azure AD Enterprise Application (service principal), _not_ the App Registration's Object ID.

### 1. Find Required IDs

- **Application (client) ID:**  
  Azure AD → App Registrations → Your App → Overview → "Application (client) ID"

- **Enterprise Application (Service Principal) Object ID:**  
  Azure AD → Enterprise Applications → Your App → Overview → "Object ID"

### 2. Confirm Service Principal Exists

```powershell
Connect-MgGraph -TenantId <your-tenant-id>
Get-MgServicePrincipal -Filter "AppId eq '<your-app-client-id>'"
```
If not present, create it:
```powershell
New-MgServicePrincipal -AppId <your-app-client-id>
```

### 3. Register the Service Principal in Exchange Online

Use **Windows PowerShell** (not PowerShell Core):

```powershell
Connect-ExchangeOnline
New-ServicePrincipal -AppId <your-app-client-id> -ServiceId <enterprise-app-object-id>
```

### 4. Add the Service Principal to "Organization Management"

```powershell
Add-RoleGroupMember -Identity "Organization Management" -Member <enterprise-app-object-id>
```
> **Tip:** The Object ID from Enterprise Applications is most reliable here.

### 5. Verify Assignment

```powershell
Get-RoleGroupMember -Identity "Organization Management"
```
Look for your app's Object ID in the results.

---

## TAP Email Template Example

Migration Manager sends users a TAP email with setup instructions:

```html
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; }
    .tap { font-size: 1.3em; font-weight: bold; color: #1976d2; }
    .instructions { margin-top: 18px; }
  </style>
</head>
<body>
  <p>Dear {FirstName} {LastName},</p>
  <p>Your Temporary Access Pass (TAP) is:</p>
  <div class="tap">{TAP}</div>
  <p>
    This TAP is <b>valid for up to {TAP_HOURS} hour(s)</b> or until first use, whichever comes first.<br>
    <i>If your organisation's policy is shorter, it will expire sooner.</i>
  </p>
  <div class="instructions">
    <strong>To set up your Passkey:</strong>
    <ol>
      <li>Go to <a href="https://myaccount.microsoft.com/">https://myaccount.microsoft.com/</a> and sign in.</li>
      <li>Go to <b>Security info</b>.</li>
      <li>Click <b>Add method</b> and choose <b>Passkey</b>.</li>
      <li>Enter your TAP code when prompted: <b>{TAP}</b></li>
      <li>Follow the instructions to set up your Passkey.</li>
    </ol>
  </div>
  <p>If you need assistance, please contact IT support.</p>
  <p>Best regards,<br/>Your IT Team</p>
</body>
</html>
```

---

## Migration Manager Workflow

1. **Assign Service Principal Permissions:**  
   Follow the steps above to ensure Migration Manager can operate using app-only authentication.

2. **Send TAP Email:**  
   Migration Manager generates and sends TAP emails to users, enabling secure initial sign-in and passkey setup.

3. **Enable User Migration:**  
   Once users have signed in and set up with TAP, Migration Manager can migrate their Exchange Online mailboxes automatically.

---

## Example Values

| Value                  | Example                                       |
|------------------------|-----------------------------------------------|
| Display Name           | Migration Manager                             |
| Application (client) ID| 95b0f7c3-1f74-423c-9b50-7e1cf5b29eb5          |
| Object ID (Enterprise) | 7dff0036-05e0-4509-bb1b-71402c76643c          |
| Directory (tenant) ID  | 74214193-01af-4cfe-9128-afdb4346dd3f          |

---

## Common Mistakes

- **Do NOT use the App Registration's Object ID** for Exchange Online permissions—use the Enterprise Application Object ID.
- Only use **Windows PowerShell** for Exchange Online service principal registration.
- Role group changes may take several minutes to propagate.

---

## References

- [App-only authentication to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps)
- [Microsoft Graph PowerShell Docs](https://aka.ms/graph/sdk/powershell/docs)
- [Temporary Access Pass in Azure AD](https://learn.microsoft.com/en-us/azure/active-directory/authentication/howto-authentication-temporary-access-pass)

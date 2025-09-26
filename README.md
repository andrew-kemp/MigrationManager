# Granting Exchange Online Admin Permissions to an Azure AD App Registration (Service Principal)

This guide explains how to allow an Azure AD app registration (service principal) to manage Exchange Online using app-only authentication.  
**Important:** For role assignments, always use the Object ID from the Enterprise Application (service principal), _not_ the App Registration Object ID.

---

## Prerequisites

- An **Azure AD app registration** created in the correct tenant.
- Azure AD and Exchange Online admin privileges.
- Latest Microsoft Graph and Exchange Online PowerShell modules.

---

## Step-by-Step Instructions

### 1. Find the IDs You Need

- **Application (client) ID:**  
  Azure AD → App Registrations → Your App → Overview → "Application (client) ID"

- **Enterprise Application (Service Principal) Object ID:**  
  Azure AD → Enterprise Applications → Your App → Overview → "Object ID"  
  > **Use this Object ID for Exchange Online role assignment—not the App Registration's Object ID!**

### 2. Confirm Service Principal Exists in Tenant

Open PowerShell and connect to Microsoft Graph:

```powershell
Connect-MgGraph -TenantId <your-tenant-id>
Get-MgServicePrincipal -Filter "AppId eq '<your-app-client-id>'"
```

- If the service principal exists, note the **Object ID** returned.
- If not, create it:
    ```powershell
    New-MgServicePrincipal -AppId <your-app-client-id>
    ```

### 3. Register the Service Principal in Exchange Online

Use **Windows PowerShell** (not PowerShell Core):

```powershell
Connect-ExchangeOnline
New-ServicePrincipal -AppId <your-app-client-id> -ServiceId <enterprise-app-object-id>
```

### 4. Add the Service Principal to "Organization Management" Role Group

```powershell
Add-RoleGroupMember -Identity "Organization Management" -Member <enterprise-app-object-id>
```

> You can also use the Application (client) ID as the member, but the **Enterprise Application Object ID** is preferred and most reliable.

### 5. Verify Assignment

```powershell
Get-RoleGroupMember -Identity "Organization Management"
```
Look for your app's Object ID in the results.

---

## Troubleshooting

- **Always use the Object ID from Azure AD → Enterprise Applications** for Exchange Online role group membership.
- If you get a “ServicePrincipalNotFound” error, confirm both IDs are correct and the service principal exists in your tenant.
- Make sure you are connected to the correct tenant.
- Role assignments may take a few minutes to propagate.

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

- **Do NOT use the App Registration's Object ID** for Exchange Online role assignment—use the Enterprise App Object ID.
- Make sure you are using **Windows PowerShell** and the latest ExchangeOnlineManagement module.
- If adding the member with `Add-RoleGroupMember` fails, double-check the Object ID and ensure the SP exists.

---

## References

- [App-only authentication to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps)
- [Microsoft Graph PowerShell Docs](https://aka.ms/graph/sdk/powershell/docs)

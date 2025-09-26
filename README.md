# EXO/Graph Service Principal Auth Tool

This PowerShell WinForms GUI helps you manage user TAP (Temporary Access Pass) creation, group membership, and notification for Microsoft Entra ID (formerly Azure AD) and Exchange Online. It uses a Service Principal (App Registration) with certificate authentication for secure, automated, app-only access.

---

## Features

- Connects to Microsoft Graph and Exchange Online with a Service Principal and certificate (PFX)
- Allows bulk user management for TAP onboarding
- Adds users to a specified group before TAP is created
- Generates Temporary Access Passes (TAP) for users
- Sends customized email notifications to users with their TAP
- Visual, interactive status and log reporting

---

## Requirements

### 1. Microsoft Entra ID App Registration

Create an App Registration in Microsoft Entra ID for the tool to use.

- Go to **Microsoft Entra admin center > Applications > App registrations > New registration**
- Name: `EXO Migration Manager` (or similar)
- Supported account types: **Single tenant**
- Redirect URI: *(leave blank - not needed for app-only)*

**After creation, note:**
- **Application (client) ID**
- **Directory (tenant) ID**

---

### 2. Certificate Authentication

- Generate a self-signed certificate or use an existing one.
- Upload the **public key (.cer)** to the App Registration (**Certificates & secrets > Certificates**).
- The tool requires the **private key (.pfx)** and password.

---

### 3. API Permissions

#### Microsoft Graph - Application Permissions

Grant these permissions (admin consent required):

| Permission Name                          | Type        | Description                           |
|------------------------------------------|-------------|---------------------------------------|
| GroupMember.ReadWrite.All                | Application | Read and write all group memberships  |
| User.Read.All                            | Application | Read all users’ full profiles         |
| UserAuthenticationMethod.ReadWrite.All   | Application | Read and write all users' auth methods|
| Mail.Send                                | Application | Send mail as any user                 |
| Policy.Read.All                          | Application | Read your organization's policies     |

#### Exchange Online - Application Permissions

Grant these permissions (admin consent required):

| Permission Name         | Type        | Description                           |
|------------------------|-------------|---------------------------------------|
| Exchange.ManageAsApp   | Application | Manage Exchange as Application        |
| Mailbox.Migration      | Application | Move mailboxes between organizations  |

- After adding, **click "Grant admin consent"** for your tenant.

---

### 4. Role Assignments

#### Add App to "Organization Management" Role Group

**Important:**  
When adding your application to Exchange Online's "Organization Management" role group, use the **Object ID from the Microsoft Entra Enterprise Application** (service principal), _not_ the App Registration's Object ID.

- Go to **Exchange Admin Center > Permissions > Admin Roles**
- Edit the **Organization Management** role group
- Add your **App Registration (Enterprise Application)** as a member  
  *(Use the Object ID from Microsoft Entra admin center > Enterprise Applications > [Your App] > Overview)*

Or, use PowerShell (recommended):

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Add the Enterprise Application Object ID to Organization Management
Add-RoleGroupMember -Identity "Organization Management" -Member <enterprise-app-object-id>
```

---

### 5. Conditional Access

- Ensure Conditional Access policies do **not** block app-only authentication for the App Registration.
- Exclude the App from policies requiring MFA, compliant device, or blocking service principals.

---

## Usage

1. **Fill in your Tenant Name, App ID, Tenant ID, Certificate Path, and TAP Group ObjectID in the GUI.**
2. **Connect** - Establishes session with Exchange Online and Microsoft Graph.
3. **Upload or enter user emails for batch processing.**
4. **Send TAP** - The app will:
    - Confirm group ObjectID (visually with log)
    - Add users to the group (if not already members)
    - Create TAP for each user
    - Send customized email with the TAP to each user

---

## Permissions Reference (Summary Table)

| Scope/Role Group              | Value Needed                              | Where to Find             |
|-------------------------------|-------------------------------------------|---------------------------|
| App Registration (client) ID  | Application (client) ID                   | Microsoft Entra ID > App Registrations |
| Service Principal Object ID   | Object ID (Enterprise Application)        | Microsoft Entra ID > Enterprise Applications |
| Organization Management Role  | Add Service Principal Object ID as member | Exchange Admin Center or PowerShell |

---

## Troubleshooting

- Ensure all API permissions are **Application** type and **admin consented**.
- Conditional Access policies must allow app-only/service principal access.
- The App must be an **owner** of the group if group membership management is restricted.
- For dynamic or hybrid groups, membership changes via Graph may not be possible.
- Check the log window in the app for real-time errors and diagnostics.

---

## Security

- Keep your certificate files (.pfx) and passwords secure.
- Rotate certificates periodically and remove unused ones from Microsoft Entra ID.

---

## Licensing

Licensed under the MIT License.

---

## Author

Andrew Kemp  
[GitHub: andrew-kemp](https://github.com/andrew-kemp)

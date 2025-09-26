# EXO/Graph Service Principal Auth Tool

This PowerShell WinForms GUI helps you manage user TAP (Temporary Access Pass) creation, group membership, and notification for Microsoft Entra ID (Azure AD) and Exchange Online. It uses a Service Principal (App Registration) with certificate authentication for secure, automated, app-only access.

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

### 1. Azure AD App Registration

Create an App Registration in Azure AD for the tool to use.

- Go to **Azure Portal > Azure Active Directory > App registrations > New registration**
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

#### Add App to "Organization Management" Role

- Go to **Exchange Admin Center > Permissions > Admin Roles**
- Edit the **Organization Management** role
- Add your **App Registration (Service Principal)** as a member

#### Add App as Owner to TAP Group

- Go to **Azure AD > Groups > [Your TAP Group]**
- Under **Owners**, add your **App Registration**
- This is essential if group membership is restricted to owners

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

## Troubleshooting

- Ensure all API permissions are **Application** type and **admin consented**.
- Conditional Access policies must allow app-only/service principal access.
- The App must be an **owner** of the group if group membership management is restricted.
- For dynamic or hybrid groups, membership changes via Graph may not be possible.
- Check the log window in the app for real-time errors and diagnostics.

---

## Security

- Keep your certificate files (.pfx) and passwords secure.
- Rotate certificates periodically and remove unused ones from Azure.

---

## Licensing

Licensed under the MIT License.

---

## Author

Andrew Kemp  
[GitHub: andrew-kemp](https://github.com/andrew-kemp)# MigrationManager

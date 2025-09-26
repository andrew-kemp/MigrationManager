# EXO/Graph Service Principal Auth Tool

This PowerShell WinForms GUI helps you manage user TAP (Temporary Access Pass) creation, group membership, and notification for Microsoft Entra ID and Exchange Online. It uses a Service Principal (App Registration) with certificate authentication for secure, automated, app-only access.

---

## Features

- Connects to Microsoft Graph and Exchange Online with a Service Principal and certificate (PFX)
- Allows bulk user management for TAP onboarding
- Adds users to a specified group before TAP is created
- Generates Temporary Access Passes (TAP) for users
- Sends customized email notifications to users with their TAP
- Visual, interactive status and log reporting
- Migration Manager: Enables mailbox migration and onboarding flows after TAP is used

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

#### PowerShell Steps

1. **Find your App IDs in Microsoft Entra admin center:**
   - **Application (client) ID** from App Registrations
   - **Enterprise Application Object ID** from Enterprise Applications

2. **Register the Service Principal in Exchange Online:**
   - You must supply both the **AppId** and the **ServiceId (Enterprise App Object ID)**:
     ```powershell
     Connect-ExchangeOnline
     New-ServicePrincipal -AppId <application-client-id> -ServiceId <enterprise-app-object-id>
     ```
     Example:
     ```powershell
     New-ServicePrincipal -AppId 95b0f7c3-1f74-423c-9b50-7e1cf5b29eb5 -ServiceId 011afdf8-84cb-44cb-af6c-08a71e7c3bdf
     ```

3. **Add the Service Principal to Organization Management:**
    ```powershell
    Add-RoleGroupMember -Identity "Organization Management" -Member <enterprise-app-object-id>
    ```

---

### 5. Conditional Access

- Ensure Conditional Access policies do **not** block app-only authentication for the App Registration.
- Exclude the App from policies requiring MFA, compliant device, or blocking service principals.

---

## Migration Manager Workflow

1. **Assign Service Principal Permissions:**  
   Follow the steps above to ensure Migration Manager can operate using app-only authentication.

2. **Send TAP Email:**  
   Migration Manager generates and sends TAP emails to users, enabling secure initial sign-in and passkey setup.

3. **Enable User Migration:**  
   Once users have signed in and set up with TAP, Migration Manager can migrate their Exchange Online mailboxes automatically.

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

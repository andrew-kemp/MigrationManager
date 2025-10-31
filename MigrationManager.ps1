# --- SECTION 1: MAIN GUI FORM & INI LOGIC ---

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Write-Ini {
    param([string]$File, [hashtable]$Values)
    $sb = New-Object System.Text.StringBuilder
    $sb.AppendLine("[defaults]") | Out-Null
    foreach ($k in $Values.Keys) {
        if ($k -notin @(
            "last_batch",
            "last_csv_path",
            "tap_group_objectid",
            "mfa_group_objectid",
            "license_group_objectid",
            "intune_office_group_objectid",
            "SmtpServer",
            "SmtpFrom",
            "SmtpDisplayName",
            "ContactPersonName",
            "ContactPersonEmail"
        )) {
            $sb.AppendLine("$k=$($Values[$k])") | Out-Null
        }
    }
    foreach ($special in @(
        "last_batch",
        "last_csv_path",
        "tap_group_objectid",
        "mfa_group_objectid",
        "license_group_objectid",
        "intune_office_group_objectid",
        "SmtpServer",
        "SmtpFrom",
        "SmtpDisplayName",
        "ContactPersonName",
        "ContactPersonEmail"
    )) {
        if ($Values[$special]) {
            $sb.AppendLine("$special=$($Values[$special])") | Out-Null
        }
    }
    [IO.File]::WriteAllText($File, $sb.ToString())
}

function Read-Ini {
    param([string]$File)
    $result = @{}
    if (-not (Test-Path $File)) { return $result }
    foreach ($line in Get-Content $File) {
        if ($line -match "^\s*([a-zA-Z0-9_]+)\s*=\s*(.*)$") {
            $k = $matches[1]
            $v = $matches[2]
            $result[$k] = $v
        }
    }
    return $result
}

function Ensure-Module {
    param(
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $false)][System.Windows.Forms.TextBox]$StatusBox
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        if ($StatusBox) { $StatusBox.AppendText("Module $($ModuleName) not found. Installing...`r`n") }
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            if ($StatusBox) { $StatusBox.AppendText("$($ModuleName) installed.`r`n") }
        } catch {
            $msg = $_.Exception.Message
            if ($StatusBox) { $StatusBox.AppendText("Failed to install $($ModuleName): $msg`r`n") }
            else { Write-Host "Failed to install $($ModuleName): $msg" }
            return $false
        }
    } else {
        if ($StatusBox) { $StatusBox.AppendText("$($ModuleName) is already installed.`r`n") }
    }
    return $true
}

function Test-EXOMailbox {
    param([string]$Email)
    try {
        $mbx = Get-Mailbox -Identity $Email -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

$iniFile = Join-Path -Path (Get-Location) -ChildPath "exo_graph_auth_gui.ini"
$logFile = Join-Path -Path (Get-Location) -ChildPath "migration_log.txt"

$defaults = Read-Ini $iniFile
if (-not $defaults["TenantName"])   { $defaults["TenantName"]   = "kempy" }
if (-not $defaults["AppId"])        { $defaults["AppId"]        = "" }
if (-not $defaults["TenantId"])     { $defaults["TenantId"]     = "" }
if (-not $defaults["CertPath"])     { $defaults["CertPath"]     = "" }
if (-not $defaults["tap_group_objectid"]) { $defaults["tap_group_objectid"] = "" }
if (-not $defaults["mfa_group_objectid"]) { $defaults["mfa_group_objectid"] = "" }
if (-not $defaults["license_group_objectid"]) { $defaults["license_group_objectid"] = "" }
if (-not $defaults["intune_office_group_objectid"]) { $defaults["intune_office_group_objectid"] = "" }
if (-not $defaults["SmtpServer"])   { $defaults["SmtpServer"]   = "" }
if (-not $defaults["SmtpFrom"])     { $defaults["SmtpFrom"]     = "" }
if (-not $defaults["SmtpDisplayName"]) { $defaults["SmtpDisplayName"] = "" }
if (-not $defaults["ContactPersonName"]) { $defaults["ContactPersonName"] = "" }
if (-not $defaults["ContactPersonEmail"]) { $defaults["ContactPersonEmail"] = "" }

$form = New-Object System.Windows.Forms.Form
$form.Text = "EXO/Graph/On-Prem Migration Manager"
$form.Size = New-Object System.Drawing.Size(700,800)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- Certificate Panel ---
$certGroup = New-Object System.Windows.Forms.GroupBox
$certGroup.Text = "Certificate for Authentication"
$certGroup.Size = New-Object System.Drawing.Size(650,160)
$certGroup.Location = New-Object System.Drawing.Point(15,10)

$rbExisting = New-Object System.Windows.Forms.RadioButton
$rbExisting.Text = "Use existing PFX Certificate"
$rbExisting.Location = New-Object System.Drawing.Point(20,25)
$rbExisting.Size = New-Object System.Drawing.Size(200,20)
$rbExisting.Checked = $true

$rbCreate = New-Object System.Windows.Forms.RadioButton
$rbCreate.Text = "Create new Certificate"
$rbCreate.Location = New-Object System.Drawing.Point(320,25)
$rbCreate.Size = New-Object System.Drawing.Size(200,20)

$panelExisting = New-Object System.Windows.Forms.Panel
$panelExisting.Size = New-Object System.Drawing.Size(610,50)
$panelExisting.Location = New-Object System.Drawing.Point(20,50)

$txtPfxPath = New-Object System.Windows.Forms.TextBox
$txtPfxPath.Location = New-Object System.Drawing.Point(0,5)
$txtPfxPath.Size = New-Object System.Drawing.Size(350,20)
$txtPfxPath.ReadOnly = $true
$txtPfxPath.Text = $defaults["CertPath"]

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Location = New-Object System.Drawing.Point(360,3)
$btnBrowse.Size = New-Object System.Drawing.Size(70,22)
$btnBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "PFX files (*.pfx)|*.pfx"
    if($ofd.ShowDialog() -eq "OK") {
        $txtPfxPath.Text = $ofd.FileName
    }
})

$lblPfxPwd = New-Object System.Windows.Forms.Label
$lblPfxPwd.Text = "Password:"
$lblPfxPwd.Location = New-Object System.Drawing.Point(0,30)
$lblPfxPwd.Size = New-Object System.Drawing.Size(60,20)

$txtPfxPwd = New-Object System.Windows.Forms.TextBox
$txtPfxPwd.Location = New-Object System.Drawing.Point(70,28)
$txtPfxPwd.Size = New-Object System.Drawing.Size(120,20)
$txtPfxPwd.UseSystemPasswordChar = $true

$panelExisting.Controls.AddRange(@($txtPfxPath,$btnBrowse,$lblPfxPwd,$txtPfxPwd))

$panelCreate = New-Object System.Windows.Forms.Panel
$panelCreate.Size = New-Object System.Drawing.Size(610,90)
$panelCreate.Location = New-Object System.Drawing.Point(20,50)
$panelCreate.Visible = $false

$lblSaveTo = New-Object System.Windows.Forms.Label
$lblSaveTo.Text = "Save As (PFX):"
$lblSaveTo.Location = New-Object System.Drawing.Point(0,5)
$lblSaveTo.Size = New-Object System.Drawing.Size(85,20)

$txtSavePath = New-Object System.Windows.Forms.TextBox
$txtSavePath.Location = New-Object System.Drawing.Point(90,3)
$txtSavePath.Size = New-Object System.Drawing.Size(220,20)
$txtSavePath.ReadOnly = $true

$btnSaveBrowse = New-Object System.Windows.Forms.Button
$btnSaveBrowse.Text = "..."
$btnSaveBrowse.Location = New-Object System.Drawing.Point(320,1)
$btnSaveBrowse.Size = New-Object System.Drawing.Size(30,22)
$btnSaveBrowse.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "PFX files (*.pfx)|*.pfx"
    if($sfd.ShowDialog() -eq "OK") {
        $txtSavePath.Text = $sfd.FileName
    }
})

$lblNewPwd = New-Object System.Windows.Forms.Label
$lblNewPwd.Text = "Password:"
$lblNewPwd.Location = New-Object System.Drawing.Point(0,35)
$lblNewPwd.Size = New-Object System.Drawing.Size(60,20)

$txtNewPwd = New-Object System.Windows.Forms.TextBox
$txtNewPwd.Location = New-Object System.Drawing.Point(70,33)
$txtNewPwd.Size = New-Object System.Drawing.Size(100,20)
$txtNewPwd.UseSystemPasswordChar = $true

$lblConfirmPwd = New-Object System.Windows.Forms.Label
$lblConfirmPwd.Text = "Confirm:"
$lblConfirmPwd.Location = New-Object System.Drawing.Point(180,35)
$lblConfirmPwd.Size = New-Object System.Drawing.Size(55,20)

$txtConfirmPwd = New-Object System.Windows.Forms.TextBox
$txtConfirmPwd.Location = New-Object System.Drawing.Point(240,33)
$txtConfirmPwd.Size = New-Object System.Drawing.Size(100,20)
$txtConfirmPwd.UseSystemPasswordChar = $true

$btnCreateCert = New-Object System.Windows.Forms.Button
$btnCreateCert.Text = "Create Certificate"
$btnCreateCert.Location = New-Object System.Drawing.Point(410,30)
$btnCreateCert.Size = New-Object System.Drawing.Size(140,28)

$panelCreate.Controls.AddRange(@(
    $lblSaveTo,$txtSavePath,$btnSaveBrowse,
    $lblNewPwd,$txtNewPwd,$lblConfirmPwd,$txtConfirmPwd,$btnCreateCert
))

$certGroup.Controls.AddRange(@($rbExisting,$rbCreate,$panelExisting,$panelCreate))

$rbExisting.Add_CheckedChanged({
    $panelExisting.Visible = $rbExisting.Checked
    $panelCreate.Visible = $rbCreate.Checked
})
$rbCreate.Add_CheckedChanged({
    $panelExisting.Visible = $rbExisting.Checked
    $panelCreate.Visible = $rbCreate.Checked
})

$btnCreateCert.Add_Click({
    $pfxPath = $txtSavePath.Text.Trim()
    $pwd = $txtNewPwd.Text
    $pwd2 = $txtConfirmPwd.Text
    if (!$pfxPath) {
        [System.Windows.Forms.MessageBox]::Show("Please choose where to save the PFX certificate file.")
        return
    }
    if ($pwd -ne $pwd2 -or !$pwd) {
        [System.Windows.Forms.MessageBox]::Show("Passwords do not match or are empty.")
        return
    }
    try {
        $securePwd = ConvertTo-SecureString $pwd -AsPlainText -Force
        $cert = New-SelfSignedCertificate -Subject "CN=EXOAppCert" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature
        Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePwd
        $cerPath = [System.IO.Path]::ChangeExtension($pfxPath, ".cer")
        Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
        [System.Windows.Forms.MessageBox]::Show("Certificate created:`n$pfxPath`n$cerPath`nUpload the .cer file to your Azure AD App Registration, and use the .pfx for authentication.")
        $rbExisting.Checked = $true
        $rbCreate.Checked = $false
        $txtPfxPath.Text = $pfxPath
        $txtPfxPwd.Text = ""
        $txtNewPwd.Text = ""
        $txtConfirmPwd.Text = ""
        $txtSavePath.Text = ""
        $defaults["CertPath"] = $pfxPath
        Write-Ini $iniFile $defaults
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to create certificate files: $($_.Exception.Message)")
    }
})

# --- Connection Settings ---
$lblTenantName = New-Object System.Windows.Forms.Label
$lblTenantName.Text = "Tenant Name:"
$lblTenantName.Location = New-Object System.Drawing.Point(30,185)
$lblTenantName.Size = New-Object System.Drawing.Size(90,20)

$txtTenantName = New-Object System.Windows.Forms.TextBox
$txtTenantName.Location = New-Object System.Drawing.Point(120,183)
$txtTenantName.Size = New-Object System.Drawing.Size(260,20)
$txtTenantName.Text = $defaults["TenantName"]

$lblAppId = New-Object System.Windows.Forms.Label
$lblAppId.Text = "App (Client) ID:"
$lblAppId.Location = New-Object System.Drawing.Point(30,215)
$lblAppId.Size = New-Object System.Drawing.Size(90,20)

$txtAppId = New-Object System.Windows.Forms.TextBox
$txtAppId.Location = New-Object System.Drawing.Point(120,213)
$txtAppId.Size = New-Object System.Drawing.Size(260,20)
$txtAppId.Text = $defaults["AppId"]

$lblTenantId = New-Object System.Windows.Forms.Label
$lblTenantId.Text = "Tenant ID (GUID):"
$lblTenantId.Location = New-Object System.Drawing.Point(30,245)
$lblTenantId.Size = New-Object System.Drawing.Size(110,20)

$txtTenantId = New-Object System.Windows.Forms.TextBox
$txtTenantId.Location = New-Object System.Drawing.Point(140,243)
$txtTenantId.Size = New-Object System.Drawing.Size(240,20)
$txtTenantId.Text = $defaults["TenantId"]

$lblTapGroup = New-Object System.Windows.Forms.Label
$lblTapGroup.Text = "TAP Group ObjectID:"
$lblTapGroup.Location = New-Object System.Drawing.Point(30,275)
$lblTapGroup.Size = New-Object System.Drawing.Size(130,20)

$txtTapGroup = New-Object System.Windows.Forms.TextBox
$txtTapGroup.Location = New-Object System.Drawing.Point(160,273)
$txtTapGroup.Size = New-Object System.Drawing.Size(300,20)
$txtTapGroup.Text = $defaults["tap_group_objectid"]

$lblMfaGroup = New-Object System.Windows.Forms.Label
$lblMfaGroup.Text = "Regular MFA Group ObjectID:"
$lblMfaGroup.Location = New-Object System.Drawing.Point(30,305)
$lblMfaGroup.Size = New-Object System.Drawing.Size(180,20)
$txtMfaGroup = New-Object System.Windows.Forms.TextBox
$txtMfaGroup.Location = New-Object System.Drawing.Point(210,303)
$txtMfaGroup.Size = New-Object System.Drawing.Size(250,20)
$txtMfaGroup.Text = $defaults["mfa_group_objectid"]

$lblLicenseGroup = New-Object System.Windows.Forms.Label
$lblLicenseGroup.Text = "License Group ObjectID:"
$lblLicenseGroup.Location = New-Object System.Drawing.Point(30,335)
$lblLicenseGroup.Size = New-Object System.Drawing.Size(150,20)

$txtLicenseGroup = New-Object System.Windows.Forms.TextBox
$txtLicenseGroup.Location = New-Object System.Drawing.Point(180,333)
$txtLicenseGroup.Size = New-Object System.Drawing.Size(280,20)
$txtLicenseGroup.Text = $defaults["license_group_objectid"]

$lblOfficeGroup = New-Object System.Windows.Forms.Label
$lblOfficeGroup.Text = "Intune Office Deployment Group GUID:"
$lblOfficeGroup.Location = New-Object System.Drawing.Point(30,365)
$lblOfficeGroup.Size = New-Object System.Drawing.Size(250,20)

$txtOfficeGroup = New-Object System.Windows.Forms.TextBox
$txtOfficeGroup.Location = New-Object System.Drawing.Point(280,363)
$txtOfficeGroup.Size = New-Object System.Drawing.Size(250,20)
$txtOfficeGroup.Text = $defaults["intune_office_group_objectid"]

$lblSmtpServer = New-Object System.Windows.Forms.Label
$lblSmtpServer.Text = "SMTP Server:"
$lblSmtpServer.Location = New-Object System.Drawing.Point(30,395)
$lblSmtpServer.Size = New-Object System.Drawing.Size(100,20)

$txtSmtpServer = New-Object System.Windows.Forms.TextBox
$txtSmtpServer.Location = New-Object System.Drawing.Point(140,393)
$txtSmtpServer.Size = New-Object System.Drawing.Size(260,20)
$txtSmtpServer.Text = $defaults["SmtpServer"]

$lblSmtpFrom = New-Object System.Windows.Forms.Label
$lblSmtpFrom.Text = "SMTP From Address:"
$lblSmtpFrom.Location = New-Object System.Drawing.Point(30,425)
$lblSmtpFrom.Size = New-Object System.Drawing.Size(120,20)

$txtSmtpFrom = New-Object System.Windows.Forms.TextBox
$txtSmtpFrom.Location = New-Object System.Drawing.Point(160,423)
$txtSmtpFrom.Size = New-Object System.Drawing.Size(250,20)
$txtSmtpFrom.Text = $defaults["SmtpFrom"]

$lblSmtpDisplay = New-Object System.Windows.Forms.Label
$lblSmtpDisplay.Text = "Sender Display Name:"
$lblSmtpDisplay.Location = New-Object System.Drawing.Point(30,455)
$lblSmtpDisplay.Size = New-Object System.Drawing.Size(130,20)

$txtSmtpDisplay = New-Object System.Windows.Forms.TextBox
$txtSmtpDisplay.Location = New-Object System.Drawing.Point(170,453)
$txtSmtpDisplay.Size = New-Object System.Drawing.Size(240,20)
$txtSmtpDisplay.Text = $defaults["SmtpDisplayName"]

# --- Contact Person Section ---
$lblContactName = New-Object System.Windows.Forms.Label
$lblContactName.Text = "Contact Person Name:"
$lblContactName.Location = New-Object System.Drawing.Point(30,485)
$lblContactName.Size = New-Object System.Drawing.Size(130,20)

$txtContactName = New-Object System.Windows.Forms.TextBox
$txtContactName.Location = New-Object System.Drawing.Point(170,483)
$txtContactName.Size = New-Object System.Drawing.Size(240,20)
$txtContactName.Text = $defaults["ContactPersonName"]

$lblContactEmail = New-Object System.Windows.Forms.Label
$lblContactEmail.Text = "Contact Person Email:"
$lblContactEmail.Location = New-Object System.Drawing.Point(30,515)
$lblContactEmail.Size = New-Object System.Drawing.Size(130,20)

$txtContactEmail = New-Object System.Windows.Forms.TextBox
$txtContactEmail.Location = New-Object System.Drawing.Point(170,513)
$txtContactEmail.Size = New-Object System.Drawing.Size(240,20)
$txtContactEmail.Text = $defaults["ContactPersonEmail"]

$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Location = New-Object System.Drawing.Point(20,550)
$txtStatus.Size = New-Object System.Drawing.Size(650,120)
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = "Vertical"
$txtStatus.ReadOnly = $true

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect"
$btnConnect.Location = New-Object System.Drawing.Point(250,700)
$btnConnect.Size = New-Object System.Drawing.Size(180,38)

$form.Controls.AddRange(@(
    $certGroup,
    $lblTenantName,$txtTenantName,
    $lblAppId,$txtAppId,
    $lblTenantId,$txtTenantId,
    $lblTapGroup,$txtTapGroup,
    $lblMfaGroup,$txtMfaGroup,
    $lblLicenseGroup,$txtLicenseGroup,
    $lblOfficeGroup,$txtOfficeGroup,
    $lblSmtpServer,$txtSmtpServer,
    $lblSmtpFrom,$txtSmtpFrom,
    $lblSmtpDisplay,$txtSmtpDisplay,
    $lblContactName, $txtContactName,
    $lblContactEmail, $txtContactEmail,
    $btnConnect,
    $txtStatus
))

$form.AcceptButton = $btnConnect
$txtPfxPath.Add_TextChanged({ $defaults["CertPath"] = $txtPfxPath.Text; Write-Ini $iniFile $defaults })
$txtTenantName.Add_TextChanged({ $defaults["TenantName"] = $txtTenantName.Text; Write-Ini $iniFile $defaults })
$txtAppId.Add_TextChanged({ $defaults["AppId"] = $txtAppId.Text; Write-Ini $iniFile $defaults })
$txtTenantId.Add_TextChanged({ $defaults["TenantId"] = $txtTenantId.Text; Write-Ini $iniFile $defaults })
$txtTapGroup.Add_TextChanged({ $defaults["tap_group_objectid"] = $txtTapGroup.Text; Write-Ini $iniFile $defaults })
$txtMfaGroup.Add_TextChanged({ $defaults["mfa_group_objectid"] = $txtMfaGroup.Text; Write-Ini $iniFile $defaults })
$txtLicenseGroup.Add_TextChanged({ $defaults["license_group_objectid"] = $txtLicenseGroup.Text; Write-Ini $iniFile $defaults })
$txtOfficeGroup.Add_TextChanged({ $defaults["intune_office_group_objectid"] = $txtOfficeGroup.Text; Write-Ini $iniFile $defaults })
$txtSmtpServer.Add_TextChanged({ $defaults["SmtpServer"] = $txtSmtpServer.Text; Write-Ini $iniFile $defaults })
$txtSmtpFrom.Add_TextChanged({ $defaults["SmtpFrom"] = $txtSmtpFrom.Text; Write-Ini $iniFile $defaults })
$txtSmtpDisplay.Add_TextChanged({ $defaults["SmtpDisplayName"] = $txtSmtpDisplay.Text; Write-Ini $iniFile $defaults })
$txtContactName.Add_TextChanged({ $defaults["ContactPersonName"] = $txtContactName.Text; Write-Ini $iniFile $defaults })
$txtContactEmail.Add_TextChanged({ $defaults["ContactPersonEmail"] = $txtContactEmail.Text; Write-Ini $iniFile $defaults })

function Send-SmtpMail {
    param(
        [string]$To,
        [string]$Subject,
        [string]$BodyHtml,
        [string]$SmtpServer,
        [string]$From,
        [string]$DisplayName
    )
    try {
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = New-Object System.Net.Mail.MailAddress($From, $DisplayName)
        $msg.To.Add($To)
        $msg.Subject = $Subject
        $msg.Body = $BodyHtml
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer)
        $smtp.Send($msg)
        return $true
    } catch {
        return $_.Exception.Message
    }
}

$btnConnect.Add_Click({
    $txtStatus.Text = ""
    $tenantName = $txtTenantName.Text.Trim()
    $appId     = $txtAppId.Text.Trim()
    $tenantId  = $txtTenantId.Text.Trim()
    $tapGroup  = $txtTapGroup.Text.Trim()
    $mfaGroup  = $txtMfaGroup.Text.Trim()
    $licenseGroup = $txtLicenseGroup.Text.Trim()
    $officeGroup = $txtOfficeGroup.Text.Trim()
    $smtpServer = $txtSmtpServer.Text.Trim()
    $smtpFrom = $txtSmtpFrom.Text.Trim()
    $smtpDisplay = $txtSmtpDisplay.Text.Trim()
    $contactName = $txtContactName.Text.Trim()
    $contactEmail = $txtContactEmail.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($tenantName) -or [string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($tenantId)) {
        $txtStatus.Text = "Tenant Name, App ID, and Tenant ID are required."
        return
    }
    $defaults["tap_group_objectid"] = $tapGroup
    $defaults["mfa_group_objectid"] = $mfaGroup
    $defaults["license_group_objectid"] = $licenseGroup
    $defaults["intune_office_group_objectid"] = $officeGroup
    $defaults["SmtpServer"] = $smtpServer
    $defaults["SmtpFrom"] = $smtpFrom
    $defaults["SmtpDisplayName"] = $smtpDisplay
    $defaults["ContactPersonName"] = $contactName
    $defaults["ContactPersonEmail"] = $contactEmail
    Write-Ini $iniFile $defaults
    $exoOrg = "$tenantName.onmicrosoft.com"
    $certPath = $txtPfxPath.Text.Trim()
    $certPwd = $txtPfxPwd.Text
    if (-not (Test-Path $certPath)) {
        $txtStatus.Text = "Please select a valid PFX file."
        return
    }
    if ([string]::IsNullOrWhiteSpace($certPwd)) {
        $txtStatus.Text = "Certificate password cannot be empty."
        return
    }
    $exoOk = Ensure-Module -ModuleName "ExchangeOnlineManagement" -StatusBox $txtStatus
    $graphOk = Ensure-Module -ModuleName "Microsoft.Graph" -StatusBox $txtStatus
    if (-not $exoOk -or -not $graphOk) {
        $txtStatus.AppendText("Cannot continue unless all modules are installed.")
        return
    }
    $securePwd = ConvertTo-SecureString $certPwd -AsPlainText -Force
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath, $securePwd)
    } catch {
        $txtStatus.Text = "Certificate error: $($_.Exception.Message)"
        return
    }
    $txtStatus.AppendText("Connecting to Exchange Online...`r`n")
    $exoOK = $false
    try {
        Connect-ExchangeOnline -AppId $appId -Organization $exoOrg -CertificateFilePath $certPath -CertificatePassword $securePwd -ShowBanner:$false
        $txtStatus.AppendText("Connected to Exchange Online!`r`n")
        $exoOK = $true
    } catch {
        $txtStatus.AppendText("EXO connection error: $($_.Exception.Message)`r`n")
    }
    $txtStatus.AppendText("Connecting to Microsoft Graph (App-Only)...`r`n")
    $graphOK = $false
    try {
        Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $cert.Thumbprint
        $txtStatus.AppendText("Connected to Graph!`r`n")
        $graphOK = $true
    } catch {
        $txtStatus.AppendText("Graph connection error: $($_.Exception.Message)`r`n")
    }
    if ($exoOK -and $graphOK) {
        $txtStatus.AppendText("Connection successful. Opening batch form...`r`n")
        $form.Hide()
        Show-BatchForm $iniFile $logFile $defaults $null
        $form.Show()
    }
})

######### Section 2: BATCH PROCESSING FORM & LOGIC #########
function Show-BatchForm {
    param($iniFile, $logFile, $defaults, $batchToLoad)

    # --- Helper: Ensure batch name exists ---
    function EnsureBatchName {
        if (-not $txtBatchName.Text.Trim()) {
            $now = Get-Date -Format "yyyy-MM-dd_HH-mm"
            $txtBatchName.Text = "MigrationBatch_$now"
            $defaults["last_batch"] = $txtBatchName.Text
            Write-Ini $iniFile $defaults
        }
        return $txtBatchName.Text.Trim()
    }

    # --- Helper: Refresh ListView with batch status/info ---
    function Refresh-ListView {
        param($batchName)
        $lstUsers.Items.Clear()
        $users = @{}
        if ((Test-Path $logFile) -and $batchName) {
            foreach ($line in Get-Content $logFile) {
                $parts = $line -split ","
                if ($parts.Count -ge 9 -and $parts[1] -eq $batchName) {
                    $users[$parts[2]] = $true
                }
            }
        }
        $tapGroupId     = $defaults["tap_group_objectid"]
        $licenseGroupId = $defaults["license_group_objectid"]
        $officeGroupId  = $defaults["intune_office_group_objectid"]

        foreach ($email in $users.Keys) {
            $userObj = $null
            try { $userObj = Get-MgUser -UserId $email -ErrorAction Stop }
            catch { try { $userObj = Get-MgUser -Filter "mail eq '$email'" -ErrorAction Stop } catch { $userObj = $null } }
            $userId = $null
            if ($userObj) { $userId = $userObj.Id }

            $addedToGroup = "No"
            if ($tapGroupId -and $userId -and (Is-GroupMember $tapGroupId $userId)) { $addedToGroup = "Yes" }
            $licensed = "No"
            if ($licenseGroupId -and $userId -and (Is-GroupMember $licenseGroupId $userId)) { $licensed = "Yes" }
            $officeDeployed = "No"
            if ($officeGroupId -and $userId -and (Is-GroupMember $officeGroupId $userId)) { $officeDeployed = "Yes" }
            $tapSet  = $addedToGroup
            $tapSent = $addedToGroup

            $migrationStatus = "Not Started"
            $mailStatusIsCompleted = $false
            try {
                $moveReq = Get-MoveRequest -Identity $email -ErrorAction Stop
                if ($moveReq.Status) { $migrationStatus = $moveReq.Status }
                if ($migrationStatus -eq "Completed") { $mailStatusIsCompleted = $true }
            } catch {
                try {
                    $mbx = Get-Mailbox -Identity $email -ErrorAction Stop
                    if ($mbx.RecipientTypeDetails -like "*MailUser*") {
                        $migrationStatus = "Not Started"
                    } else {
                        $migrationStatus = "Completed"
                        $mailStatusIsCompleted = $true
                    }
                } catch {
                    $migrationStatus = "Not Started"
                }
            }

            $status = "Not Started"
            if ($mailStatusIsCompleted -or $migrationStatus -eq "Completed") {
                $status = "Completed"
            } elseif (
                $tapSent -eq "Yes" -or
                $licensed -eq "Yes" -or
                $officeDeployed -eq "Yes" -or
                ($migrationStatus -ne "Not Started" -and $migrationStatus -ne "" -and $migrationStatus -ne $null -and $migrationStatus -ne "Completed")
            ) {
                $status = "In Progress"
            }

            $item = New-Object System.Windows.Forms.ListViewItem ($email)
            $item.SubItems.Add($status) | Out-Null
            $item.SubItems.Add($addedToGroup) | Out-Null
            $item.SubItems.Add($tapSet) | Out-Null
            $item.SubItems.Add($tapSent) | Out-Null
            $item.SubItems.Add($licensed) | Out-Null
            $item.SubItems.Add($officeDeployed) | Out-Null
            $item.SubItems.Add($migrationStatus) | Out-Null
            $lstUsers.Items.Add($item)
        }
    }

    # --- Main Window & Controls ---
    $batchForm = New-Object System.Windows.Forms.Form
    $batchForm.Text = "User Migration Batch"
    $batchForm.Size = New-Object System.Drawing.Size(1120, 1000)
    $batchForm.StartPosition = "CenterScreen"
    $batchForm.FormBorderStyle = "FixedDialog"
    $batchForm.MaximizeBox = $false

    $lblEndpoint = New-Object System.Windows.Forms.Label
    $lblEndpoint.Text = "Migration Endpoint:"
    $lblEndpoint.Location = New-Object System.Drawing.Point(20, 20)
    $lblEndpoint.Size = New-Object System.Drawing.Size(125, 20)
    $cmbEndpoint = New-Object System.Windows.Forms.ComboBox
    $cmbEndpoint.Location = New-Object System.Drawing.Point(150, 18)
    $cmbEndpoint.Size = New-Object System.Drawing.Size(400, 22)
    $cmbEndpoint.DropDownStyle = "DropDownList"
    try {
        $endpoints = Get-MigrationEndpoint | Select-Object -ExpandProperty Identity
        foreach ($ep in $endpoints) { $cmbEndpoint.Items.Add($ep) }
        if ($cmbEndpoint.Items.Count -gt 0) { $cmbEndpoint.SelectedIndex = 0 }
    } catch {
        $cmbEndpoint.Items.Add("No endpoints available")
        $cmbEndpoint.SelectedIndex = 0
    }

    $btnRemoveUser = New-Object System.Windows.Forms.Button
    $btnRemoveUser.Text = "Remove Selected"
    $btnRemoveUser.Location = New-Object System.Drawing.Point(600, 18)
    $btnRemoveUser.Size = New-Object System.Drawing.Size(130, 28)

    $btnClearBatch = New-Object System.Windows.Forms.Button
    $btnClearBatch.Text = "Clear Batch"
    $btnClearBatch.Location = New-Object System.Drawing.Point(740, 18)
    $btnClearBatch.Size = New-Object System.Drawing.Size(130, 28)

    $btnNewBatch = New-Object System.Windows.Forms.Button
    $btnNewBatch.Text = "New Batch"
    $btnNewBatch.Location = New-Object System.Drawing.Point(880, 18)
    $btnNewBatch.Size = New-Object System.Drawing.Size(90, 28)

    $btnUploadCSV = New-Object System.Windows.Forms.Button
    $btnUploadCSV.Text = "Upload CSV"
    $btnUploadCSV.Location = New-Object System.Drawing.Point(620, 60)
    $btnUploadCSV.Size = New-Object System.Drawing.Size(130, 28)

    $btnAddUsers = New-Object System.Windows.Forms.Button
    $btnAddUsers.Text = "Add Users"
    $btnAddUsers.Location = New-Object System.Drawing.Point(776, 60)
    $btnAddUsers.Size = New-Object System.Drawing.Size(130, 28)

    $btnRefreshList = New-Object System.Windows.Forms.Button
    $btnRefreshList.Text = "Refresh List"
    $btnRefreshList.Location = New-Object System.Drawing.Point(20, 60)
    $btnRefreshList.Size = New-Object System.Drawing.Size(120, 28)

    $lblUsers = New-Object System.Windows.Forms.Label
    $lblUsers.Text = "Users in Batch (Email, Status, Added to Group, TAP Set, TAP Sent, Licensed, Office Deployed, Mailbox Migration Status):"
    $lblUsers.Location = New-Object System.Drawing.Point(20, 90)
    $lblUsers.Size = New-Object System.Drawing.Size(1080, 20)

    $lstUsers = New-Object System.Windows.Forms.ListView
    $lstUsers.Location = New-Object System.Drawing.Point(20, 120)
    $lstUsers.Size = New-Object System.Drawing.Size(1070, 430)
    $lstUsers.View = [System.Windows.Forms.View]::Details
    $lstUsers.Columns.Add("Email", 240) | Out-Null
    $lstUsers.Columns.Add("Status", 100) | Out-Null
    $lstUsers.Columns.Add("Added to Group", 100) | Out-Null
    $lstUsers.Columns.Add("TAP Set", 80) | Out-Null
    $lstUsers.Columns.Add("TAP Sent", 80) | Out-Null
    $lstUsers.Columns.Add("Licensed", 80) | Out-Null
    $lstUsers.Columns.Add("Office Deployed", 100) | Out-Null
    $lstUsers.Columns.Add("Mailbox Migration Status", 170) | Out-Null
    $lstUsers.FullRowSelect = $true
    $lstUsers.HideSelection = $false
    $lstUsers.GridLines = $true

    # --- Batch Loader Panel ---
    $panelBatchLoader = New-Object System.Windows.Forms.Panel
    $panelBatchLoader.Location = New-Object System.Drawing.Point(20, 570)
    $panelBatchLoader.Size = New-Object System.Drawing.Size(540, 40)

    $lblChooseBatch = New-Object System.Windows.Forms.Label
    $lblChooseBatch.Text = "Load Existing Batch:"
    $lblChooseBatch.Location = New-Object System.Drawing.Point(0, 10)
    $lblChooseBatch.Size = New-Object System.Drawing.Size(130, 20)

    $cmbBatches = New-Object System.Windows.Forms.ComboBox
    $cmbBatches.Location = New-Object System.Drawing.Point(140, 8)
    $cmbBatches.Size = New-Object System.Drawing.Size(250, 22)
    $cmbBatches.DropDownStyle = "DropDownList"
    $cmbBatches.Items.Clear()

    # Build a mapping of batch technical names to friendly names
    $batchNameToFriendly = @{}
    foreach ($key in $defaults.Keys) {
        if ($key -like "batch_friendlyname_*") {
            $batchId = $key.Substring(19)
            $batchNameToFriendly[$batchId] = $defaults[$key]
        }
    }
    $allBatchNames = @(Get-Content $logFile | ForEach-Object { ($_ -split ",")[1] }) | Select-Object -Unique
    foreach ($b in $allBatchNames) {
        $friendly = $batchNameToFriendly[$b]
        if ($b) {
            if ($friendly) {
                $cmbBatches.Items.Add("$b ($friendly)")
            } else {
                $cmbBatches.Items.Add($b)
            }
        }
    }

    $btnLoadBatch = New-Object System.Windows.Forms.Button
    $btnLoadBatch.Text = "Load"
    $btnLoadBatch.Location = New-Object System.Drawing.Point(400, 6)
    $btnLoadBatch.Size = New-Object System.Drawing.Size(100, 24)

    $panelBatchLoader.Controls.AddRange(@($lblChooseBatch, $cmbBatches, $btnLoadBatch))

    $lblBatch = New-Object System.Windows.Forms.Label
    $lblBatch.Text = "Batch Name:"
    $lblBatch.Location = New-Object System.Drawing.Point(20, 620)
    $lblBatch.Size = New-Object System.Drawing.Size(120, 20)
    $txtBatchName = New-Object System.Windows.Forms.TextBox
    $txtBatchName.Location = New-Object System.Drawing.Point(140, 617)
    $txtBatchName.Size = New-Object System.Drawing.Size(350, 24)
    $txtBatchName.ReadOnly = $false

    $btnGenBatch = New-Object System.Windows.Forms.Button
    $btnGenBatch.Text = "Generate"
    $btnGenBatch.Location = New-Object System.Drawing.Point(510, 617)
    $btnGenBatch.Size = New-Object System.Drawing.Size(100, 24)
    $btnGenBatch.Add_Click({
        $now = Get-Date -Format "yyyy-MM-dd_HH-mm"
        $txtBatchName.Text = "MigrationBatch_$now"
    })

    # --- Batch Friendly Name ---
    $lblBatchFriendly = New-Object System.Windows.Forms.Label
    $lblBatchFriendly.Text = "Batch Friendly Name:"
    $lblBatchFriendly.Location = New-Object System.Drawing.Point(20, 660)
    $lblBatchFriendly.Size = New-Object System.Drawing.Size(150, 20)

    $txtBatchFriendly = New-Object System.Windows.Forms.TextBox
    $txtBatchFriendly.Location = New-Object System.Drawing.Point(180, 657)
    $txtBatchFriendly.Size = New-Object System.Drawing.Size(310, 24)

    $btnRenameBatch = New-Object System.Windows.Forms.Button
    $btnRenameBatch.Text = "Rename Batch"
    $btnRenameBatch.Location = New-Object System.Drawing.Point(510, 657)
    $btnRenameBatch.Size = New-Object System.Drawing.Size(100, 24)
    $btnRenameBatch.Add_Click({
        $oldBatchName = $txtBatchName.Text.Trim()
        $newFriendlyName = $txtBatchFriendly.Text.Trim()
        if ($oldBatchName -and $newFriendlyName) {
            # Update friendly name mapping
            $defaults["batch_friendlyname_$oldBatchName"] = $newFriendlyName
            Write-Ini $iniFile $defaults

            # Update batch dropdown with new friendly name
            $cmbBatches.Items.Clear()
            $batchNameToFriendly = @{}
            foreach ($key in $defaults.Keys) {
                if ($key -like "batch_friendlyname_*") {
                    $batchId = $key.Substring(19)
                    $batchNameToFriendly[$batchId] = $defaults[$key]
                }
            }
            $allBatchNames = @(Get-Content $logFile | ForEach-Object { ($_ -split ",")[1] }) | Select-Object -Unique
            foreach ($b in $allBatchNames) {
                $friendly = $batchNameToFriendly[$b]
                if ($b) {
                    if ($friendly) {
                        $cmbBatches.Items.Add("$b ($friendly)")
                    } else {
                        $cmbBatches.Items.Add($b)
                    }
                }
            }
            $lblSaved.Text = "Batch $oldBatchName renamed to: $newFriendlyName"
        }
    })

    # --- Auth Method GroupBox ---
    $gbAuthMethod = New-Object System.Windows.Forms.GroupBox
    $gbAuthMethod.Text = "Authentication Method"
    $gbAuthMethod.Location = New-Object System.Drawing.Point(20, 700)
    $gbAuthMethod.Size = New-Object System.Drawing.Size(340, 60)

    $rbTAP = New-Object System.Windows.Forms.RadioButton
    $rbTAP.Text = "Passkey (TAP/Passkey)"
    $rbTAP.Location = New-Object System.Drawing.Point(15, 25)
    $rbTAP.Size = New-Object System.Drawing.Size(150, 22)

    $rbMFA = New-Object System.Windows.Forms.RadioButton
    $rbMFA.Text = "Regular MFA"
    $rbMFA.Location = New-Object System.Drawing.Point(180, 25)
    $rbMFA.Size = New-Object System.Drawing.Size(120, 22)
    $rbMFA.Checked = $true      # Default

    $gbAuthMethod.Controls.AddRange(@($rbTAP, $rbMFA))

    # --- Device Entry for Office Deployment ---
    $lblDeviceEntry = New-Object System.Windows.Forms.Label
    $lblDeviceEntry.Text = "Device Name for Office Deployment:"
    $lblDeviceEntry.Location = New-Object System.Drawing.Point(400, 700)
    $lblDeviceEntry.Size = New-Object System.Drawing.Size(240, 20)

    $txtDeviceEntry = New-Object System.Windows.Forms.TextBox
    $txtDeviceEntry.Location = New-Object System.Drawing.Point(400, 730)
    $txtDeviceEntry.Size = New-Object System.Drawing.Size(340, 24)

    $btnDeployOffice = New-Object System.Windows.Forms.Button
    $btnDeployOffice.Text = "Deploy Office to Device"
    $btnDeployOffice.Location = New-Object System.Drawing.Point(760, 730)
    $btnDeployOffice.Size = New-Object System.Drawing.Size(180, 28)

    # --- Migration Steps GroupBox ---
    $gbMigrationSteps = New-Object System.Windows.Forms.GroupBox
    $gbMigrationSteps.Text = "Migration Steps"
    $gbMigrationSteps.Location = New-Object System.Drawing.Point(20, 770)
    $gbMigrationSteps.Size = New-Object System.Drawing.Size(600, 60)

    $chkSendTAP = New-Object System.Windows.Forms.CheckBox
    $chkSendTAP.Text = "Send TAP"
    $chkSendTAP.Location = New-Object System.Drawing.Point(20, 25)
    $chkSendTAP.Size = New-Object System.Drawing.Size(110, 22)
    $chkSendTAP.Checked = $false

    $chkLicenseUser = New-Object System.Windows.Forms.CheckBox
    $chkLicenseUser.Text = "License User"
    $chkLicenseUser.Location = New-Object System.Drawing.Point(140, 25)
    $chkLicenseUser.Size = New-Object System.Drawing.Size(110, 22)
    $chkLicenseUser.Checked = $false

    $chkMigrateMailbox = New-Object System.Windows.Forms.CheckBox
    $chkMigrateMailbox.Text = "Migrate Mailbox"
    $chkMigrateMailbox.Location = New-Object System.Drawing.Point(260, 25)
    $chkMigrateMailbox.Size = New-Object System.Drawing.Size(120, 22)
    $chkMigrateMailbox.Checked = $false

    $chkCompleteMailbox = New-Object System.Windows.Forms.CheckBox
    $chkCompleteMailbox.Text = "Complete Mailbox"
    $chkCompleteMailbox.Location = New-Object System.Drawing.Point(390, 25)
    $chkCompleteMailbox.Size = New-Object System.Drawing.Size(130, 22)
    $chkCompleteMailbox.Checked = $false

    $gbMigrationSteps.Controls.AddRange(@($chkSendTAP, $chkLicenseUser, $chkMigrateMailbox, $chkCompleteMailbox))

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "Apply Selected Tasks"
    $btnApply.Location = New-Object System.Drawing.Point(650, 790)
    $btnApply.Size = New-Object System.Drawing.Size(150, 32)
    $btnResendTAP = New-Object System.Windows.Forms.Button
    $btnResendTAP.Text = "Re-Send TAP"
    $btnResendTAP.Location = New-Object System.Drawing.Point(820, 790)
    $btnResendTAP.Size = New-Object System.Drawing.Size(120, 32)

    $lblSaved = New-Object System.Windows.Forms.Label
    $lblSaved.Text = ""
    $lblSaved.Location = New-Object System.Drawing.Point(20, 840)
    $lblSaved.Size = New-Object System.Drawing.Size(950, 20)

    # --- Enable Send TAP only when Passkey is selected ---
    $chkSendTAP.Enabled = $rbTAP.Checked
    $rbTAP.Add_CheckedChanged({
        $chkSendTAP.Enabled = $rbTAP.Checked
    })
    $rbMFA.Add_CheckedChanged({
        $chkSendTAP.Enabled = $rbTAP.Checked
    })

    $batchForm.Controls.AddRange(@(
        $lblEndpoint, $cmbEndpoint,
        $btnRemoveUser, $btnClearBatch, $btnNewBatch, $btnUploadCSV, $btnAddUsers, $btnRefreshList,
        $lblUsers, $lstUsers,
        $panelBatchLoader,
        $lblBatch, $txtBatchName, $btnGenBatch,
        $lblBatchFriendly, $txtBatchFriendly, $btnRenameBatch,
        $gbAuthMethod,
        $lblDeviceEntry, $txtDeviceEntry, $btnDeployOffice,
        $gbMigrationSteps,
        $btnApply, $btnResendTAP,
        $lblSaved
    ))


    # --- Device Deploy Office Button Handler ---
    $btnDeployOffice.Add_Click({
        $officeGroupId = $defaults["intune_office_group_objectid"]
        $deviceName = $txtDeviceEntry.Text.Trim()
        if (-not $deviceName) {
            [System.Windows.Forms.MessageBox]::Show("Enter a device name to deploy Office.")
            return
        }
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Office Deployment Progress"
        $progressForm.Size = New-Object System.Drawing.Size(650,300)
        $txtProgress = New-Object System.Windows.Forms.TextBox
        $txtProgress.Multiline = $true
        $txtProgress.ScrollBars = "Vertical"
        $txtProgress.ReadOnly = $true
        $txtProgress.Dock = "Fill"
        $progressForm.Controls.Add($txtProgress)
        $progressForm.Show()

        # Find devices by name, sort by lastCheckInDateTime (newest first)
        try {
            $devices = Get-MgDevice -All | Where-Object { $_.displayName -like "*$deviceName*" }
        } catch { $devices = @() }
        if (-not $devices -or $devices.Count -eq 0) {
            $txtProgress.AppendText("No device found for '$deviceName'.`r`n")
            return
        }
        $sorted = $devices | Sort-Object -Property lastCheckInDateTime -Descending
        $newest = $sorted | Select-Object -First 1
        $oldDevices = $sorted | Select-Object -Skip 1

        $msg = "Devices found for '$deviceName':`r`n"
        $idx = 1
        foreach ($dev in $sorted) {
            $msg += "$idx. ObjectId: $($dev.Id) - LastCheckIn: $($dev.lastCheckInDateTime)`r`n"
            $idx++
        }
        $msg += "`r`nAdd the newest device to Office group and remove others from Entra/Intune?"
        $ans = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Office Deployment", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($ans -ne "Yes") {
            $txtProgress.AppendText("Cancelled by user.`r`n")
            return
        }

        # Add newest to Office group
        try {
            New-MgGroupMember -GroupId $officeGroupId -DirectoryObjectId $newest.Id -ErrorAction Stop
            $txtProgress.AppendText("Added device '$deviceName' (ObjectId: $($newest.Id)) to Office group.`r`n")
        } catch {
            $txtProgress.AppendText("Error adding device '$deviceName': $($_.Exception.Message)`r`n")
        }

        # Remove older devices
        if ($oldDevices.Count -gt 0) {
            foreach ($dev in $oldDevices) {
                try {
                    Remove-MgDevice -DeviceId $dev.Id -ErrorAction Stop
                    $txtProgress.AppendText("Deleted device ObjectId: $($dev.Id)`r`n")
                } catch {
                    $txtProgress.AppendText("Error deleting device ObjectId: $($dev.Id) - $($_.Exception.Message)`r`n")
                }
            }
        } else {
            $txtProgress.AppendText("No older devices to remove.`r`n")
        }
        $txtProgress.AppendText("Office deployment step finished for device '$deviceName'.`r`n")
    })

    # --- Upload CSV Handler ---
    $btnUploadCSV.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "CSV files (*.csv)|*.csv"
        $ofd.Title = "Select CSV file with user emails"
        if ($ofd.ShowDialog() -eq "OK") {
            $csvPath = $ofd.FileName
            $curBatch = EnsureBatchName
            $added = 0
            $users = @{}
            if ((Test-Path $logFile) -and $curBatch) {
                foreach ($line in Get-Content $logFile) {
                    $parts = $line -split ","
                    if ($parts.Count -ge 9 -and $parts[1] -eq $curBatch) {
                        $users[$parts[2]] = $true
                    }
                }
            }
            $csvContent = Import-Csv -Path $csvPath
            foreach ($row in $csvContent) {
                $email = $row.Email
                if (-not $email) {
                    $email = $row.PSObject.Properties[0].Value
                }
                $email = $email.Trim()
                if ($email -and $email -match "@" -and -not $users.ContainsKey($email)) {
                    $dt = Get-Date -Format "yyyy-MM-dd HH:mm"
                    $line = "$dt,$curBatch,$email,Not Started,No,No,No,No,No,Not Started"
                    Add-Content -Path $logFile -Value $line
                    $added++
                }
            }
            Refresh-ListView $curBatch
            $lblSaved.Text = "$added user(s) added to batch $curBatch from CSV."
        }
    })

    # --- Remove Selected Users Handler ---
    $btnRemoveUser.Add_Click({
        $curBatch = $txtBatchName.Text.Trim()
        if ($curBatch) {
            $selected = @($lstUsers.SelectedItems)
            if ($selected.Count -gt 0 -and (Test-Path $logFile)) {
                $lines = Get-Content $logFile
                foreach ($sel in $selected) {
                    $lines = $lines | Where-Object { -not ($_ -match "^[^,]+,${curBatch},$($sel.Text),") }
                }
                Set-Content $logFile $lines
                Refresh-ListView $curBatch
                $lblSaved.Text = "$($selected.Count) user(s) removed from batch ${curBatch}."
            }
        }
    })

    # --- Clear Batch Handler ---
    $btnClearBatch.Add_Click({
        $batchName = $txtBatchName.Text.Trim()
        if ($batchName -and (Test-Path $logFile)) {
            $lines = Get-Content $logFile | Where-Object { -not ($_ -match "^[^,]+,${batchName},") }
            Set-Content $logFile $lines
            Refresh-ListView $batchName
            $lblSaved.Text = "All users removed from batch ${batchName}."
        }
    })

    # --- Refresh List Handler ---
    $btnRefreshList.Add_Click({
        $curBatch = $txtBatchName.Text.Trim()
        if ($curBatch) {
            Refresh-ListView $curBatch
            $lblSaved.Text = "User status refreshed."
        }
    })

    # --- New Batch Handler ---
    $btnNewBatch.Add_Click({
        $now = Get-Date -Format "yyyy-MM-dd_HH-mm"
        $batchName = "MigrationBatch_$now"
        $txtBatchName.Text = $batchName
        $defaults["last_batch"] = $batchName
        Write-Ini $iniFile $defaults
        $lstUsers.Items.Clear()
        $lblSaved.Text = "New batch created: $batchName. Add users or upload CSV."
    })

    # --- Load Batch Handler ---
    $btnLoadBatch.Add_Click({
        $selectedBatch = $cmbBatches.SelectedItem
        if ($selectedBatch) {
            $txtBatchName.Text = $selectedBatch
            Refresh-ListView $selectedBatch
            $defaults["last_batch"] = $selectedBatch
            Write-Ini $iniFile $defaults
            $lblSaved.Text = "Batch ${selectedBatch} loaded."
        }
    })

    # --- Add User Button Handler ---
    $btnAddUsers.Add_Click({
        $addForm = New-Object System.Windows.Forms.Form
        $addForm.Text = "Add Users"
        $addForm.Size = New-Object System.Drawing.Size(400,300)
        $addForm.StartPosition = "CenterParent"
        $lblAdd = New-Object System.Windows.Forms.Label
        $lblAdd.Text = "Enter email addresses (comma or line separated):"
        $lblAdd.Location = New-Object System.Drawing.Point(10,10)
        $lblAdd.Size = New-Object System.Drawing.Size(370,20)
        $txtAdd = New-Object System.Windows.Forms.TextBox
        $txtAdd.Location = New-Object System.Drawing.Point(10,40)
        $txtAdd.Size = New-Object System.Drawing.Size(360,160)
        $txtAdd.Multiline = $true
        $txtAdd.ScrollBars = "Vertical"
        $btnAdd = New-Object System.Windows.Forms.Button
        $btnAdd.Text = "Add"
        $btnAdd.Location = New-Object System.Drawing.Point(150,220)
        $btnAdd.Size = New-Object System.Drawing.Size(90,32)
        $lblAdded = New-Object System.Windows.Forms.Label
        $lblAdded.Text = ""
        $lblAdded.Location = New-Object System.Drawing.Point(10,200)
        $lblAdded.Size = New-Object System.Drawing.Size(360,20)
        $btnAdd.Add_Click({
            $block = $txtAdd.Text
            $curBatch = EnsureBatchName
            $added = 0
            $users = @{}
            if ((Test-Path $logFile) -and $curBatch) {
                foreach ($line in Get-Content $logFile) {
                    $parts = $line -split ","
                    if ($parts.Count -ge 9 -and $parts[1] -eq $curBatch) {
                        $users[$parts[2]] = $true
                    }
                }
            }
            if ($block -and $curBatch) {
                $emails = $block -split "[,`n`r]" | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -match "@" }
                foreach ($e in $emails) {
                    if (-not $users.ContainsKey($e)) {
                        $dt = Get-Date -Format "yyyy-MM-dd HH:mm"
                        $line = "$dt,$curBatch,$e,Not Started,No,No,No,No,No,Not Started"
                        Add-Content -Path $logFile -Value $line
                        $added++
                    }
                }
                Refresh-ListView $curBatch
                $txtAdd.Text = ""
                $lblAdded.Text = "$added user(s) added to batch $curBatch."
            }
        })
        $addForm.Controls.AddRange(@($lblAdd,$txtAdd,$btnAdd,$lblAdded))
        [void]$addForm.ShowDialog()
    })

        # --- Apply Selected Tasks Handler ---
        $btnApply.Add_Click({
        if ($lstUsers.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select at least one user to apply actions.")
            return
        }
        $tapGroupId     = $defaults["tap_group_objectid"]
        $mfaGroupId     = $defaults["mfa_group_objectid"]
        $licenseGroupId = $defaults["license_group_objectid"]
        $smtpServer     = $defaults["SmtpServer"]
        $smtpFrom       = $defaults["SmtpFrom"]
        $smtpDisplay    = $defaults["SmtpDisplayName"]
        $templateTap    = "TAP_Email_Template.html"
        $templateMfa    = "MFA_Email_Template.html"
        $cutoverTemplate = "Mailbox_Cutover_Email_Template.html"
        $curBatch       = $txtBatchName.Text
        $tenantName     = $defaults["TenantName"]

        $useTAP = $rbTAP.Checked
        $useMFA = $rbMFA.Checked

        $endpointMap = @{}
        try {
            $endpoints = Get-MigrationEndpoint
            foreach ($ep in $endpoints) {
            $endpointMap[$ep.Identity] = $ep.RemoteServer
            }
        } catch {
            $endpointMap[$cmbEndpoint.SelectedItem] = $cmbEndpoint.SelectedItem
        }
        $selectedEndpointName = $cmbEndpoint.SelectedItem
        $remoteHostName = $endpointMap[$selectedEndpointName]
        $targetDeliveryDomain = "$tenantName.mail.onmicrosoft.com"

        if ($chkMigrateMailbox.Checked -and $chkMigrateMailbox.Enabled -and -not $Global:ExchangeCreds) {
            $Global:ExchangeCreds = Get-Credential -Message "Enter on-prem Exchange admin credentials"
        }

        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Processing Selected Tasks"
        $progressForm.Size = New-Object System.Drawing.Size(650,400)
        $txtProgress = New-Object System.Windows.Forms.TextBox
        $txtProgress.Multiline = $true
        $txtProgress.ScrollBars = "Vertical"
        $txtProgress.ReadOnly = $true
        $txtProgress.Dock = "Fill"
        $progressForm.Controls.Add($txtProgress)
        $progressForm.Show()

        # --- TAP/Passkey logic: batch group add and TAP creation ---
        if ($useTAP -and $tapGroupId) {
            # Remove selected users from MFA group if present (Cross-validation)
            foreach ($sel in $lstUsers.SelectedItems) {
            $userEmail = $sel.Text
            $userObj = $null
            try { $userObj = Get-MgUser -UserId $userEmail -ErrorAction Stop }
            catch { try { $userObj = Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop } catch { $userObj = $null } }
            if ($userObj) {
                $userId = $userObj.Id
                $inMfaGroup = $false
                try {
                $inMfaGroup = (Get-MgGroupMember -GroupId $mfaGroupId -All | Where-Object { $_.Id -eq $userId }) -ne $null
                } catch {}
                if ($inMfaGroup) {
                try {
                    Remove-MgGroupMemberByRef  -GroupId $mfaGroupId -DirectoryObjectId $userId -ErrorAction Stop
                    $txtProgress.AppendText("Removed $userEmail from MFA group due to Passkey selection.`r`n")
                } catch {
                    $txtProgress.AppendText("Error removing $userEmail from MFA group: $($_.Exception.Message)`r`n")
                }
                } else {
                $txtProgress.AppendText("$userEmail is not a member of MFA group; nothing to remove.`r`n")
                }
            }
            }
            # Auto-tick Send TAP Checkbox
            $chkSendTAP.Checked = $true

            $tapResults = @{}
            foreach ($sel in $lstUsers.SelectedItems) {
            $userEmail = $sel.Text
            $userObj = $null
            try { $userObj = Get-MgUser -UserId $userEmail -ErrorAction Stop }
            catch { try { $userObj = Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop } catch { $userObj = $null } }
            if (-not $userObj) {
                $txtProgress.AppendText("User $userEmail not found in Entra ID.`r`n")
                $tapResults[$userEmail] = $false
                continue
            }
            $userId = $userObj.Id
            $alreadyInTapGroup = $false
            try {
                $alreadyInTapGroup = (Get-MgGroupMember -GroupId $tapGroupId -All | Where-Object { $_.Id -eq $userId }) -ne $null
            } catch { $alreadyInTapGroup = $false }
            if ($alreadyInTapGroup) {
                $txtProgress.AppendText("User $userEmail is already in TAP group.`r`n")
                $tapResults[$userEmail] = $true
            } else {
                try {
                New-MgGroupMember -GroupId $tapGroupId -DirectoryObjectId $userId -ErrorAction Stop
                $txtProgress.AppendText("Added $userEmail to TAP group.`r`n")
                $tapResults[$userEmail] = $true
                } catch {
                $txtProgress.AppendText("Error adding $userEmail to TAP group: $($_.Exception.Message)`r`n")
                $tapResults[$userEmail] = $false
                }
            }
            }

            $txtProgress.AppendText("Waiting 10 seconds for TAP policy to apply to all users...`r`n")
            Start-Sleep -Seconds 10

            foreach ($sel in $lstUsers.SelectedItems) {
            $userEmail = $sel.Text
            if (-not $tapResults[$userEmail]) {
                $txtProgress.AppendText("Skipping TAP creation for $userEmail (not in TAP group).`r`n")
                continue
            }
            $userObj = $null
            try { $userObj = Get-MgUser -UserId $userEmail -ErrorAction Stop }
            catch { try { $userObj = Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop } catch { $userObj = $null } }
            if (-not $userObj) { continue }
            $userId = $userObj.Id

            # --- Check for existing valid TAP ---
            $existingTap = $null
            $tap = $null
            $tapValid = $false
            try {
                $existingTap = Get-MgUserAuthenticationTemporaryAccessPassMethod -UserId $userId -ErrorAction SilentlyContinue
            } catch { $existingTap = $null }
            if ($existingTap -and $existingTap.TemporaryAccessPass) {
                $now = Get-Date
                if ($existingTap.validityWindowEndDateTime -gt $now) {
                $tap = $existingTap.TemporaryAccessPass
                $tapValid = $true
                $txtProgress.AppendText("User $userEmail already has a valid TAP. Skipping new TAP creation and email.`r`n")
                }
            }
            if (-not $tapValid) {
                # --- Create TAP with retry ---
                $tapCreated = $false
                $maxTries = 10
                $tryCount = 0
                $waitSeconds = 20
                $policyApplies = $false
                while (-not $policyApplies -and $tryCount -lt $maxTries) {
                try {
                    $tapMethod = New-MgUserAuthenticationTemporaryAccessPassMethod `
                    -UserId $userId `
                    -BodyParameter @{lifetimeInMinutes=480; isUsableOnce=$true}
                    $tap = $tapMethod.TemporaryAccessPass
                    $tapCreated = $true
                    $policyApplies = $true
                    $txtProgress.AppendText("TAP created for $userEmail.`r`n")
                } catch {
                    if ($_.Exception.Message -like "*UserCredentialPolicy does not allow*") {
                    $txtProgress.AppendText("Waiting for TAP policy to apply to $userEmail (attempt $($tryCount+1)/$maxTries)...`r`n")
                    Start-Sleep -Seconds $waitSeconds
                    $tryCount++
                    } else {
                    $txtProgress.AppendText("Error creating TAP: $($_.Exception.Message)`r`n")
                    break
                    }
                }
                }
                if (-not $policyApplies) {
                $txtProgress.AppendText("ERROR: TAP policy did not apply after waiting. Please try again later or check policy assignments.`r`n")
                continue
                }
            }
            # --- Send TAP email if TAP was created or valid ---
            if ($tap) {
                $firstName = $userObj.GivenName
                $lastName = $userObj.Surname
                if (-not (Test-Path $templateTap)) {
                $txtProgress.AppendText("Email template not found: ${templateTap}`r`n")
                } else {
                $html = Get-Content $templateTap -Raw
                $html = $html -replace "\{FirstName\}", $firstName
                $html = $html -replace "\{LastName\}", $lastName
                $html = $html -replace "\{TAP\}", $tap
                $html = $html -replace "\{TAP_HOURS\}", 8
                $html = $html -replace "\{SenderDisplayName\}", $smtpDisplay
                $html = $html -replace "\{ContactPersonName\}", $defaults["ContactPersonName"]
                $html = $html -replace "\{ContactPersonEmail\}", $defaults["ContactPersonEmail"]

                $sent = Send-SmtpMail -To $userEmail -Subject "Your Temporary Access Pass (TAP) for Passkey Setup" -BodyHtml $html -SmtpServer $smtpServer -From $smtpFrom -DisplayName $smtpDisplay
                if ($sent -eq $true) { $txtProgress.AppendText("TAP email sent to $userEmail.`r`n") }
                else { $txtProgress.AppendText(("SMTP error for {0}: {1}`r`n" -f $userEmail, $sent)) }
                }
            }
            }
        }

        # --- MFA logic: batch removal from TAP if selected, group add/email ---
        if ($useMFA -and $mfaGroupId) {
            # Remove selected users from TAP group if present (Cross-validation)
            foreach ($sel in $lstUsers.SelectedItems) {
            $userEmail = $sel.Text
            $userObj = $null
            try { $userObj = Get-MgUser -UserId $userEmail -ErrorAction Stop }
            catch { try { $userObj = Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop } catch { $userObj = $null } }
            if ($userObj) {
                $userId = $userObj.Id
                $inTapGroup = $false
                try {
                $inTapGroup = (Get-MgGroupMember -GroupId $tapGroupId -All | Where-Object { $_.Id -eq $userId }) -ne $null
                } catch {}
                if ($inTapGroup) {
                try {
                    Remove-MgGroupMemberByRef -GroupId $tapGroupId -DirectoryObjectId $userId -ErrorAction Stop
                    $txtProgress.AppendText("Removed $userEmail from TAP group due to MFA selection.`r`n")
                } catch {
                    $txtProgress.AppendText("Error removing $userEmail from TAP group: $($_.Exception.Message)`r`n")
                }
                } else {
                $txtProgress.AppendText("$userEmail is not a member of TAP group; nothing to remove.`r`n")
                }
            }
            }

            foreach ($sel in $lstUsers.SelectedItems) {
            $userEmail = $sel.Text
            $userObj = $null
            try { $userObj = Get-MgUser -UserId $userEmail -ErrorAction Stop }
            catch { try { $userObj = Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop } catch { $userObj = $null } }
            if (-not $userObj) {
                $txtProgress.AppendText("User $userEmail not found in Entra ID.`r`n")
                continue
            }
            $userId = $userObj.Id
            $alreadyInMfaGroup = $false
            try {
                $alreadyInMfaGroup = (Get-MgGroupMember -GroupId $mfaGroupId -All | Where-Object { $_.Id -eq $userId }) -ne $null
            } catch { $alreadyInMfaGroup = $false }
            if ($alreadyInMfaGroup) {
                $txtProgress.AppendText("User $userEmail is already in Regular MFA group. Skipping MFA email.`r`n")
            } else {
                try {
                New-MgGroupMember -GroupId $mfaGroupId -DirectoryObjectId $userId -ErrorAction Stop
                $txtProgress.AppendText("Added $userEmail to Regular MFA group.`r`n")
                } catch {
                $txtProgress.AppendText("Error adding to Regular MFA group: $($_.Exception.Message)`r`n")
                }
                if (Test-Path $templateMfa) {
                $firstName = $userObj.GivenName
                $lastName = $userObj.Surname
                $html = Get-Content $templateMfa -Raw
                $html = $html -replace "\{FirstName\}", $firstName
                $html = $html -replace "\{LastName\}", $lastName
                $html = $html -replace "\{SenderDisplayName\}", $smtpDisplay
                $html = $html -replace "\{ContactPersonName\}", $defaults["ContactPersonName"]
                $html = $html -replace "\{ContactPersonEmail\}", $defaults["ContactPersonEmail"]
                $sent = Send-SmtpMail -To $userEmail -Subject "Your MFA Setup Instructions" -BodyHtml $html -SmtpServer $smtpServer -From $smtpFrom -DisplayName $smtpDisplay
                if ($sent -eq $true) { $txtProgress.AppendText("MFA email sent to $userEmail.`r`n") }
                else { $txtProgress.AppendText(("SMTP error for {0}: {1}`r`n" -f $userEmail, $sent)) }
                } else {
                $txtProgress.AppendText("MFA email template not found.`r`n")
                }
            }
            }
        }

        # --- License group logic ---
        foreach ($sel in $lstUsers.SelectedItems) {
            $userEmail = $sel.Text
            $userObj = $null
            try { $userObj = Get-MgUser -UserId $userEmail -ErrorAction Stop }
            catch { try { $userObj = Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop } catch { $userObj = $null } }
            if (-not $userObj) {
            $txtProgress.AppendText("User $userEmail not found in Entra ID.`r`n")
            continue
            }
            $userId = $userObj.Id

            if ($chkLicenseUser.Checked -and $licenseGroupId) {
            $alreadyLicensed = $false
            try {
                $alreadyLicensed = (Get-MgGroupMember -GroupId $licenseGroupId -All | Where-Object { $_.Id -eq $userId }) -ne $null
            } catch { $alreadyLicensed = $false }
            if ($alreadyLicensed) {
                $txtProgress.AppendText("User $userEmail is already in License group.`r`n")
            } else {
                try {
                New-MgGroupMember -GroupId $licenseGroupId -DirectoryObjectId $userId -ErrorAction Stop
                $txtProgress.AppendText("Added $userEmail to License group.`r`n")
                } catch {
                $txtProgress.AppendText("Error licensing ${userEmail}: $($_.Exception.Message)`r`n")
                }
            }
            }

            # --- Licensing, mailbox migration, etc. ---
            if ($chkMigrateMailbox.Checked -and $chkMigrateMailbox.Enabled) {
            try {
                New-MoveRequest -Identity $userEmail `
                -Remote `
                -RemoteHostName $remoteHostName `
                -RemoteCredential $Global:ExchangeCreds `
                -TargetDeliveryDomain $targetDeliveryDomain `
                -BatchName $curBatch `
                -SuspendWhenReadyToComplete `
                -ErrorAction Stop
                $txtProgress.AppendText("Started migration for $userEmail using endpoint '$selectedEndpointName' (`$remoteHostName`).`r`n")
            } catch {
                $txtProgress.AppendText("Error migrating ${userEmail}: $($_.Exception.Message)`r`n")
            }
            }

            # --- Complete Mailbox & Cutover Email: unconditional (no prereq check) ---
            if ($chkCompleteMailbox.Checked -and $chkCompleteMailbox.Enabled) {
            $firstName = $userObj.GivenName
            $lastName = $userObj.Surname

            # Always resolve template file to script directory
            $scriptFolder    = Split-Path -Parent $MyInvocation.MyCommand.Path
            $cutoverTemplate = Join-Path $scriptFolder "Mailbox_Cutover_Email_Template.html"

            if (![string]::IsNullOrWhiteSpace($cutoverTemplate) -and (Test-Path $cutoverTemplate)) {
                $html = Get-Content $cutoverTemplate -Raw
                $html = $html -replace "\{FirstName\}", $firstName
                $html = $html -replace "\{LastName\}", $lastName
                $html = $html -replace "\{SenderDisplayName\}", $smtpDisplay
                $html = $html -replace "\{ContactPersonName\}", $defaults["ContactPersonName"]
                $html = $html -replace "\{ContactPersonEmail\}", $defaults["ContactPersonEmail"]
                $sent = Send-SmtpMail -To $userEmail -Subject "Your Mailbox Migration: Cutover Instructions" -BodyHtml $html -SmtpServer $smtpServer -From $smtpFrom -DisplayName $smtpDisplay
                if ($sent -eq $true) {
                $txtProgress.AppendText("Cutover email sent to $userEmail.`r`n")
                } else {
                $txtProgress.AppendText(("SMTP error for {0}: {1}`r`n" -f $userEmail, $sent))
                }
            } else {
                $txtProgress.AppendText("Mailbox cutover template not found for $userEmail.`r`n")
            }
            $migrationCompleted = $false
            try {
                Resume-MoveRequest -Identity $userEmail -ErrorAction Stop -WarningAction SilentlyContinue
                $txtProgress.AppendText("Completing migration for $userEmail.`r`n")
                $migrationCompleted = $true
            } catch {
                $txtProgress.AppendText("Error completing migration for ${userEmail}: $($_.Exception.Message)`r`n")
            }
            }
        }

        $txtProgress.AppendText("All done!`r`n")
        $lblSaved.Text = "Selected tasks applied for selected users."
        Refresh-ListView $curBatch
        # Reset all task checkboxes
        $chkSendTAP.Checked         = $false
        $chkLicenseUser.Checked     = $false
        $chkMigrateMailbox.Checked  = $false
        $chkCompleteMailbox.Checked = $false
        })


    # --- Re-Send TAP Handler (unchanged) ---
$btnResendTAP.Add_Click({
    if ($lstUsers.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one user to re-send TAP email.")
        return
    }

    $tapGroupId     = $defaults["tap_group_objectid"]
    $smtpServer     = $defaults["SmtpServer"]
    $smtpFrom       = $defaults["SmtpFrom"]
    $smtpDisplay    = $defaults["SmtpDisplayName"]
    $templateTap    = "TAP_Email_Template.html"

    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Re-Sending TAP Emails"
    $progressForm.Size = New-Object System.Drawing.Size(650,300)
    $txtProgress = New-Object System.Windows.Forms.TextBox
    $txtProgress.Multiline = $true
    $txtProgress.ScrollBars = "Vertical"
    $txtProgress.ReadOnly = $true
    $txtProgress.Dock = "Fill"
    $progressForm.Controls.Add($txtProgress)
    $progressForm.Show()

    foreach ($sel in $lstUsers.SelectedItems) {
        $userEmail = $sel.Text
        $txtProgress.AppendText("Re-sending TAP to $userEmail...`r`n")
        $userObj = $null
        try { $userObj = Get-MgUser -UserId $userEmail -ErrorAction Stop }
        catch { try { $userObj = Get-MgUser -Filter "mail eq '$userEmail'" -ErrorAction Stop } catch { $userObj = $null } }
        if (-not $userObj) {
            $txtProgress.AppendText("User $userEmail not found in Entra ID.`r`n")
            continue
        }
        $userId = $userObj.Id

        # Get TAP if exists (or create a new one if needed)
        $tap = $null
        try {
            $existingTap = Get-MgUserAuthenticationTemporaryAccessPassMethod -UserId $userId -ErrorAction SilentlyContinue
            if ($existingTap -and $existingTap.TemporaryAccessPass) {
                $tap = $existingTap.TemporaryAccessPass
            }
        } catch { $tap = $null }

        if (-not $tap) {
            # If no TAP exists, create a new one
            try {
                $tapMethod = New-MgUserAuthenticationTemporaryAccessPassMethod `
                    -UserId $userId `
                    -BodyParameter @{lifetimeInMinutes=480; isUsableOnce=$true}
                $tap = $tapMethod.TemporaryAccessPass
                $txtProgress.AppendText("New TAP created for $userEmail.`r`n")
            } catch {
                $txtProgress.AppendText("Error creating TAP: $($_.Exception.Message)`r`n")
                continue
            }
        }

        $firstName = $userObj.GivenName
        $lastName = $userObj.Surname
        if (-not (Test-Path $templateTap)) {
            $txtProgress.AppendText("Email template not found: ${templateTap}`r`n")
        } else {
            $html = Get-Content $templateTap -Raw
            $html = $html -replace "\{FirstName\}", $firstName
            $html = $html -replace "\{LastName\}", $lastName
            $html = $html -replace "\{TAP\}", $tap
            $html = $html -replace "\{TAP_HOURS\}", 8
            $html = $html -replace "\{SenderDisplayName\}", $smtpDisplay
            $html = $html -replace "\{ContactPersonName\}", $defaults["ContactPersonName"]
            $html = $html -replace "\{ContactPersonEmail\}", $defaults["ContactPersonEmail"]
            $sent = Send-SmtpMail -To $userEmail -Subject "Your Temporary Access Pass (TAP) for Passkey Setup" -BodyHtml $html -SmtpServer $smtpServer -From $smtpFrom -DisplayName $smtpDisplay
            if ($sent -eq $true) { $txtProgress.AppendText("TAP email sent to $userEmail.`r`n") }
            else { $txtProgress.AppendText(("SMTP error for {0}: {1}`r`n" -f $userEmail, $sent)) }
        }
    }
    $txtProgress.AppendText("Re-send TAP finished.`r`n")
})

    [void]$batchForm.ShowDialog()
}
# --- SECTION 3: Helper Functions & Utilities ---

# Check if user is a member of a group
function Is-GroupMember($groupId, $userId) {
    try {
        return (Get-MgGroupMember -GroupId $groupId -All | Where-Object { $_.Id -eq $userId }) -ne $null
    } catch { return $false }
}

# Get all devices by name (returns all matching device objects)
function Get-DevicesByName($deviceName) {
    try {
        return Get-MgDevice -Filter "displayName eq '$deviceName'" -ErrorAction SilentlyContinue
    } catch {
        return @()
    }
}

# Add device to group (returns success/failure)
function Add-DeviceToGroup($deviceId, $groupId) {
    try {
        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $deviceId -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Remove device by ObjectID
function Remove-Device($deviceId) {
    try {
        Remove-MgDevice -DeviceId $deviceId -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Create TAP for a user and return TAP code
function Create-TAP($userId) {
    try {
        $tapMethod = New-MgUserAuthenticationTemporaryAccessPassMethod `
            -UserId $userId `
            -BodyParameter @{lifetimeInMinutes=480; isUsableOnce=$true}
        return $tapMethod.TemporaryAccessPass
    } catch {
        return $null
    }
}

# Add user to group (returns success/failure)
function Add-UserToGroup($userId, $groupId) {
    try {
        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Send email (SMTP)
function Send-Email($to, $subject, $bodyHtml, $smtpServer, $from, $displayName) {
    try {
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = New-Object System.Net.Mail.MailAddress($from, $displayName)
        $msg.To.Add($to)
        $msg.Subject = $subject
        $msg.Body = $bodyHtml
        $msg.IsBodyHtml = $true
        $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer)
        $smtp.Send($msg)
        return $true
    } catch {
        return $_.Exception.Message
    }
}

# Get mailbox migration status for a user
function Get-MailboxMigrationStatus($email) {
    try {
        $moveReq = Get-MoveRequest -Identity $email -ErrorAction Stop
        return $moveReq.Status
    } catch {
        try {
            $mbx = Get-Mailbox -Identity $email -ErrorAction Stop
            if ($mbx.RecipientTypeDetails -like "*MailUser*") {
                return "Not Started"
            } else {
                return "Completed"
            }
        } catch {
            return "Not Started"
        }
    }
}

[void]$form.ShowDialog()

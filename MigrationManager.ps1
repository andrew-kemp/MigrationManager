Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- INI & Utility Functions ---
function Write-Ini {
    param([string]$File, [hashtable]$Values)
    $sb = New-Object System.Text.StringBuilder
    $sb.AppendLine("[defaults]") | Out-Null
    foreach ($k in $Values.Keys) {
        if ($k -ne "last_batch" -and $k -ne "last_csv_path" -and $k -ne "tap_group_objectid") {
            $sb.AppendLine("$k=$($Values[$k])") | Out-Null
        }
    }
    if ($Values["last_batch"]) {
        $sb.AppendLine("last_batch=$($Values["last_batch"])") | Out-Null
    }
    if ($Values["last_csv_path"]) {
        $sb.AppendLine("last_csv_path=$($Values["last_csv_path"])") | Out-Null
    }
    if ($Values["tap_group_objectid"]) {
        $sb.AppendLine("tap_group_objectid=$($Values["tap_group_objectid"])") | Out-Null
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

function Log-AddUser {
    param([string]$BatchName, [string]$User)
    $dt = Get-Date -Format "yyyy-MM-dd HH:mm"
    $line = "$dt,$BatchName,$User,queued"
    Add-Content -Path $logFile -Value $line
}

function Get-BatchUsersFromLog {
    param([string]$logFile, [string]$batchName)
    $users = @()
    if ((Test-Path $logFile) -and $batchName) {
        foreach ($line in Get-Content $logFile) {
            $parts = $line -split ","
            if ($parts.Count -ge 4 -and $parts[1] -eq $batchName) {
                $users += @{Email=$parts[2]; Status=$parts[3]}
            }
        }
    }
    return $users
}

function Get-AllBatchNamesFromLog {
    param([string]$logFile)
    $batches = @{}
    if (Test-Path $logFile) {
        foreach ($line in Get-Content $logFile) {
            $parts = $line -split ","
            if ($parts.Count -ge 2) {
                $batches[$parts[1]] = $true
            }
        }
    }
    return $batches.Keys
}

function Check-BatchAddOrNew {
    param($lstUsers, $txtBatchName, $defaults, $iniFile)
    if ($lstUsers.Items.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This batch already contains users. Add to existing batch (Yes), create a new batch (No), or cancel?",
            "Batch Exists",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            $now = Get-Date -Format "yyyy-MM-dd_HH-mm"
            $newBatch = "MigrationBatch_$now"
            $txtBatchName.Text = $newBatch
            $defaults["last_batch"] = $newBatch
            Write-Ini $iniFile $defaults
            $lstUsers.Items.Clear()
        }
        return $result -ne [System.Windows.Forms.DialogResult]::Cancel
    }
    return $true
}

function New-SelfSignedCertPfxAndCer {
    param(
        [string]$Subject,
        [string]$BasePath,
        [System.Security.SecureString]$Password
    )
    $pfxPath = $BasePath
    $cerPath = [System.IO.Path]::ChangeExtension($BasePath, ".cer")
    $cert = New-SelfSignedCertificate -Subject $Subject -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $Password
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
    return @($cert, $pfxPath, $cerPath)
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

# --- File/Defaults ---
$iniFile = Join-Path -Path (Get-Location) -ChildPath "exo_graph_auth_gui.ini"
$logFile = Join-Path -Path (Get-Location) -ChildPath "migration_log.txt"

$defaults = Read-Ini $iniFile
if (-not $defaults["TenantName"])   { $defaults["TenantName"]   = "kempy" }
if (-not $defaults["AppId"])        { $defaults["AppId"]        = "" }
if (-not $defaults["TenantId"])     { $defaults["TenantId"]     = "" }
if (-not $defaults["CertPath"])     { $defaults["CertPath"]     = "" }
if (-not $defaults["tap_group_objectid"]) { $defaults["tap_group_objectid"] = "" }

# --- Main GUI Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "EXO/Graph Service Principal Auth Tool"
$form.Size = New-Object System.Drawing.Size(650,560)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- Certificate Panel ---
$certGroup = New-Object System.Windows.Forms.GroupBox
$certGroup.Text = "Certificate for Authentication"
$certGroup.Size = New-Object System.Drawing.Size(610,160)
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
$panelExisting.Size = New-Object System.Drawing.Size(570,50)
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
$panelCreate.Size = New-Object System.Drawing.Size(570,90)
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
        $result = New-SelfSignedCertPfxAndCer -Subject "CN=EXOAppCert" -BasePath $pfxPath -Password $securePwd
        $cerPath = $result[2]
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

$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Location = New-Object System.Drawing.Point(20,345)
$txtStatus.Size = New-Object System.Drawing.Size(590,110)
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = "Vertical"
$txtStatus.ReadOnly = $true

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect"
$btnConnect.Location = New-Object System.Drawing.Point(200,310)
$btnConnect.Size = New-Object System.Drawing.Size(120,32)

$form.Controls.AddRange(@(
    $certGroup,
    $lblTenantName,$txtTenantName,
    $lblAppId,$txtAppId,
    $lblTenantId,$txtTenantId,
    $lblTapGroup,$txtTapGroup,
    $btnConnect,
    $txtStatus
))

$txtPfxPath.Add_TextChanged({ $defaults["CertPath"]    = $txtPfxPath.Text;    Write-Ini $iniFile $defaults })
$txtTenantName.Add_TextChanged({ $defaults["TenantName"] = $txtTenantName.Text; Write-Ini $iniFile $defaults })
$txtAppId.Add_TextChanged({ $defaults["AppId"] = $txtAppId.Text; Write-Ini $iniFile $defaults })
$txtTenantId.Add_TextChanged({ $defaults["TenantId"] = $txtTenantId.Text; Write-Ini $iniFile $defaults })
$txtTapGroup.Add_TextChanged({ $defaults["tap_group_objectid"] = $txtTapGroup.Text; Write-Ini $iniFile $defaults })

# --- Batch Form & All Logic ---
function Show-BatchForm {
    param($iniFile, $logFile, $defaults, $batchToLoad)
    $batchForm = New-Object System.Windows.Forms.Form
    $batchForm.Text = "User Migration Batch"
    $batchForm.Size = New-Object System.Drawing.Size(900,800)
    $batchForm.StartPosition = "CenterScreen"
    $batchForm.FormBorderStyle = "FixedDialog"
    $batchForm.MaximizeBox = $false

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Upload a CSV of emails or enter them manually (comma or line separated):"
    $lblInfo.Location = New-Object System.Drawing.Point(20,20)
    $lblInfo.Size = New-Object System.Drawing.Size(820,20)

    $btnUploadCSV = New-Object System.Windows.Forms.Button
    $btnUploadCSV.Text = "Upload CSV"
    $btnUploadCSV.Location = New-Object System.Drawing.Point(20,50)
    $btnUploadCSV.Size = New-Object System.Drawing.Size(100,28)

    $txtManual = New-Object System.Windows.Forms.TextBox
    $txtManual.Multiline = $true
    $txtManual.Location = New-Object System.Drawing.Point(140,50)
    $txtManual.Size = New-Object System.Drawing.Size(350,60)
    $txtManual.ScrollBars = "Vertical"

    $btnAddManual = New-Object System.Windows.Forms.Button
    $btnAddManual.Text = "Add Users"
    $btnAddManual.Location = New-Object System.Drawing.Point(510,50)
    $btnAddManual.Size = New-Object System.Drawing.Size(100,28)

    $lblUsers = New-Object System.Windows.Forms.Label
    $lblUsers.Text = "Users in Batch (Email, Status, TAP Set, Added to Group, TAP Sent):"
    $lblUsers.Location = New-Object System.Drawing.Point(20,125)
    $lblUsers.Size = New-Object System.Drawing.Size(600,20)

    $lstUsers = New-Object System.Windows.Forms.ListView
    $lstUsers.Location = New-Object System.Drawing.Point(20,150)
    $lstUsers.Size = New-Object System.Drawing.Size(850,400)
    $lstUsers.View = [System.Windows.Forms.View]::Details
    $lstUsers.Columns.Add("Email",340) | Out-Null
    $lstUsers.Columns.Add("Status",120) | Out-Null
    $lstUsers.Columns.Add("TAP Set",80) | Out-Null
    $lstUsers.Columns.Add("Added to Group",110) | Out-Null
    $lstUsers.Columns.Add("TAP Sent",100) | Out-Null
    $lstUsers.FullRowSelect = $true
    $lstUsers.HideSelection = $false
    $lstUsers.GridLines = $true

    $btnRemoveUser = New-Object System.Windows.Forms.Button
    $btnRemoveUser.Text = "Remove Selected"
    $btnRemoveUser.Location = New-Object System.Drawing.Point(20,570)
    $btnRemoveUser.Size = New-Object System.Drawing.Size(160,32)

    $btnClearBatch = New-Object System.Windows.Forms.Button
    $btnClearBatch.Text = "Clear Batch"
    $btnClearBatch.Location = New-Object System.Drawing.Point(210,570)
    $btnClearBatch.Size = New-Object System.Drawing.Size(160,32)

    $btnNewBatch = New-Object System.Windows.Forms.Button
    $btnNewBatch.Text = "New Batch"
    $btnNewBatch.Location = New-Object System.Drawing.Point(400,570)
    $btnNewBatch.Size = New-Object System.Drawing.Size(160,32)

    $btnSendTAP = New-Object System.Windows.Forms.Button
    $btnSendTAP.Text = "Send TAP"
    $btnSendTAP.Location = New-Object System.Drawing.Point(600,570)
    $btnSendTAP.Size = New-Object System.Drawing.Size(160,32)

    $lblBatch = New-Object System.Windows.Forms.Label
    $lblBatch.Text = "Batch Name:"
    $lblBatch.Location = New-Object System.Drawing.Point(20,620)
    $lblBatch.Size = New-Object System.Drawing.Size(120,20)

    $txtBatchName = New-Object System.Windows.Forms.TextBox
    $txtBatchName.Location = New-Object System.Drawing.Point(140,617)
    $txtBatchName.Size = New-Object System.Drawing.Size(350,24)
    $txtBatchName.ReadOnly = $false

    $btnGenBatch = New-Object System.Windows.Forms.Button
    $btnGenBatch.Text = "Generate"
    $btnGenBatch.Location = New-Object System.Drawing.Point(510,617)
    $btnGenBatch.Size = New-Object System.Drawing.Size(100,24)
    $btnGenBatch.Add_Click({
        $now = Get-Date -Format "yyyy-MM-dd_HH-mm"
        $txtBatchName.Text = "MigrationBatch_$now"
    })

    $lblSaved = New-Object System.Windows.Forms.Label
    $lblSaved.Text = ""
    $lblSaved.Location = New-Object System.Drawing.Point(20,650)
    $lblSaved.Size = New-Object System.Drawing.Size(850,20)

    $lblChooseBatch = New-Object System.Windows.Forms.Label
    $lblChooseBatch.Text = "Load Existing Batch:"
    $lblChooseBatch.Location = New-Object System.Drawing.Point(20,690)
    $lblChooseBatch.Size = New-Object System.Drawing.Size(130,20)

    $cmbBatches = New-Object System.Windows.Forms.ComboBox
    $cmbBatches.Location = New-Object System.Drawing.Point(160,688)
    $cmbBatches.Size = New-Object System.Drawing.Size(250,22)
    $cmbBatches.DropDownStyle = "DropDownList"
    $cmbBatches.Items.Clear()
    $allBatchNames = Get-AllBatchNamesFromLog $logFile
    foreach ($b in $allBatchNames) { $cmbBatches.Items.Add($b) }
    $btnLoadBatch = New-Object System.Windows.Forms.Button
    $btnLoadBatch.Text = "Load"
    $btnLoadBatch.Location = New-Object System.Drawing.Point(420,686)
    $btnLoadBatch.Size = New-Object System.Drawing.Size(100,24)

    # --- Helper Functions for Batch Form ---
    function Refresh-ListView {
        param($batchName)
        $lstUsers.Items.Clear()
        $users = Get-BatchUsersFromLog $logFile $batchName
        foreach ($u in $users) {
            $item = New-Object System.Windows.Forms.ListViewItem ($u.Email)
            $item.SubItems.Add($u.Status) | Out-Null
            $item.SubItems.Add("") | Out-Null # TAP Set
            $item.SubItems.Add("") | Out-Null # Added to Group
            $item.SubItems.Add("") | Out-Null # TAP Sent
            if ($u.Status -eq "queued")   { $item.SubItems[1].ForeColor = [System.Drawing.Color]::Orange }
            elseif ($u.Status -eq "migrated") { $item.SubItems[1].ForeColor = [System.Drawing.Color]::Green }
            elseif ($u.Status -eq "failed")   { $item.SubItems[1].ForeColor = [System.Drawing.Color]::Red }
            $lstUsers.Items.Add($item)
        }
    }

    function EnsureBatchName {
        if (-not $txtBatchName.Text.Trim()) {
            $now = Get-Date -Format "yyyy-MM-dd_HH-mm"
            $txtBatchName.Text = "MigrationBatch_$now"
            $defaults["last_batch"] = $txtBatchName.Text
            Write-Ini $iniFile $defaults
        }
        return $txtBatchName.Text.Trim()
    }

    $batchForm.Controls.AddRange(@(
        $lblInfo, $btnUploadCSV, $txtManual, $btnAddManual,
        $lblUsers, $lstUsers, $btnRemoveUser, $btnClearBatch, $btnNewBatch, $btnSendTAP,
        $lblBatch, $txtBatchName, $btnGenBatch,
        $lblSaved, $lblChooseBatch, $cmbBatches, $btnLoadBatch
    ))

    $btnNewBatch.Add_Click({
        $now = Get-Date -Format "yyyy-MM-dd_HH-mm"
        $batchName = "MigrationBatch_$now"
        $txtBatchName.Text = $batchName
        $defaults["last_batch"] = $batchName
        Write-Ini $iniFile $defaults
        $lstUsers.Items.Clear()
        $lblSaved.Text = "New batch created: $batchName. Add users or upload CSV."
    })

    $btnUploadCSV.Add_Click({
        if (-not (Check-BatchAddOrNew $lstUsers $txtBatchName $defaults $iniFile)) { return }
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if($ofd.ShowDialog() -eq "OK") {
            $csvPath = $ofd.FileName
            $emails = @()
            $firstline = (Get-Content $csvPath -TotalCount 1)
            if ($firstline -match "email" -or $firstline -match "Email") {
                $emails = Import-Csv $csvPath | ForEach-Object { $_.Email }
            } else {
                $emails = Get-Content $csvPath | ForEach-Object { $_.Trim() }
            }
            $emails = $emails | Where-Object { $_ -and $_ -match "@" }
            $curBatch = EnsureBatchName
            $added = 0
            foreach ($e in $emails) {
                $already = $false
                foreach ($item in $lstUsers.Items) {
                    if ($item.Text -eq $e) { $already = $true }
                }
                if (-not $already -and $curBatch) {
                    Log-AddUser $curBatch $e
                    $added++
                }
            }
            if ($curBatch) { Refresh-ListView $curBatch }
            $defaults["last_csv_path"] = $csvPath
            Write-Ini $iniFile $defaults
            $lblSaved.Text = "$added user(s) added from CSV to batch $curBatch. CSV: $csvPath"
        }
    })

    $btnAddManual.Add_Click({
        if (-not (Check-BatchAddOrNew $lstUsers $txtBatchName $defaults $iniFile)) { return }
        $block = $txtManual.Text
        $curBatch = EnsureBatchName
        $added = 0
        if ($block -and $curBatch) {
            $emails = $block -split "[,`n`r]" | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -match "@" }
            foreach ($e in $emails) {
                $already = $false
                foreach ($item in $lstUsers.Items) {
                    if ($item.Text -eq $e) { $already = $true }
                }
                if (-not $already) {
                    Log-AddUser $curBatch $e
                    $added++
                }
            }
            Refresh-ListView $curBatch
            $txtManual.Text = ""
            $lblSaved.Text = "$added user(s) added to batch $curBatch."
        }
    })

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

    $btnClearBatch.Add_Click({
        $batchName = $txtBatchName.Text.Trim()
        if ($batchName -and (Test-Path $logFile)) {
            $lines = Get-Content $logFile | Where-Object { -not ($_ -match "^[^,]+,${batchName},") }
            Set-Content $logFile $lines
            Refresh-ListView $batchName
            $lblSaved.Text = "All users removed from batch ${batchName}."
        }
    })

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

    $btnSendTAP.Add_Click({
        if ($lstUsers.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one user to send a TAP.")
            return
        }

        # --- GROUP CONFIRMATION DIALOG ---
        $inputBox = New-Object System.Windows.Forms.Form
        $inputBox.Text = "Confirm Passkey Group ObjectID"
        $inputBox.Size = New-Object System.Drawing.Size(500, 320)

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "Group ObjectID or Name:"
        $lbl.Location = New-Object System.Drawing.Point(10, 20)
        $lbl.Size = New-Object System.Drawing.Size(160, 20)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Size = New-Object System.Drawing.Size(280, 24)
        $txt.Location = New-Object System.Drawing.Point(170, 18)
        $txt.Text = $defaults["tap_group_objectid"]

        $okBtn = New-Object System.Windows.Forms.Button
        $okBtn.Text = "Confirm"
        $okBtn.Location = New-Object System.Drawing.Point(120, 60)

        $closeBtn = New-Object System.Windows.Forms.Button
        $closeBtn.Text = "Close"
        $closeBtn.Location = New-Object System.Drawing.Point(220, 60)

        $txtGroupLog = New-Object System.Windows.Forms.TextBox
        $txtGroupLog.Location = New-Object System.Drawing.Point(20, 100)
        $txtGroupLog.Size = New-Object System.Drawing.Size(440, 170)
        $txtGroupLog.Multiline = $true
        $txtGroupLog.ScrollBars = "Vertical"
        $txtGroupLog.ReadOnly = $true

        $inputBox.Controls.AddRange(@($lbl, $txt, $okBtn, $closeBtn, $txtGroupLog))

        $ok = $false
        $script:confirmedGroupObjectId = $null

        $okBtn.Add_Click({
            $txtGroupLog.Clear()
            $groupInput = $txt.Text.Trim()
            $txtGroupLog.AppendText("Looking up group: $groupInput`r`n")
            try {
                $groupObj = $null
                # Try ObjectId first
                $groupObj = Get-MgGroup -GroupId $groupInput -ErrorAction SilentlyContinue
                if (-not $groupObj) {
                    $groups = Get-MgGroup -Filter "displayName eq '$groupInput'"
                    if ($groups.Count -eq 1) {
                        $groupObj = $groups[0]
                    } elseif ($groups.Count -gt 1) {
                        $txtGroupLog.AppendText("Multiple groups found with name '$groupInput'. Please use ObjectId.`r`n")
                        return
                    }
                }
                if ($groupObj) {
                    $txtGroupLog.AppendText("Group Found:`r`n")
                    $txtGroupLog.AppendText("Display Name: $($groupObj.DisplayName)`r`n")
                    $txtGroupLog.AppendText("ObjectId: $($groupObj.Id)`r`n")
                    $txtGroupLog.AppendText("Type: $($groupObj.GroupTypes -join ', ')`r`n")
                    $txtGroupLog.AppendText("Mail: $($groupObj.Mail)`r`n")
                    if ($groupObj.MembershipRule) {
                        $txtGroupLog.AppendText("MembershipRule: $($groupObj.MembershipRule)`r`n")
                    }
                    $txtGroupLog.AppendText("`r`n")
                    $script:confirmedGroupObjectId = $groupObj.Id
                    $ok = $true
                    Start-Sleep -Seconds 10
                    $inputBox.Close()
                } else {
                    $txtGroupLog.AppendText("Group '$groupInput' not found!`r`n")
                }
            } catch {
                $txtGroupLog.AppendText("Error retrieving group info: $($_.Exception.Message)`r`n")
            }
        })

        $closeBtn.Add_Click({ $inputBox.Close() })

        $inputBox.ShowDialog()
        if (-not $ok) { return }

        $groupObjectId = $script:confirmedGroupObjectId
        $defaults["tap_group_objectid"] = $groupObjectId
        Write-Ini $iniFile $defaults

        # --- Confirmation dialog before proceeding ---
        $userList = ($lstUsers.SelectedItems | ForEach-Object { $_.Text }) -join "`r`n"
        $tapHours = 24
        $maxTAP = 1440
        try {
            $policy = Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId "TemporaryAccessPass"
            if ($policy.maximumLifetimeInMinutes) {
                $maxTAP = [math]::Min($maxTAP, [int]$policy.maximumLifetimeInMinutes)
            }
        } catch {
            $maxTAP = 1440
        }
        if ($maxTAP -eq 0) { $maxTAP = 60 }
        $tapHours = [math]::Round($maxTAP/60,0)
        $msg = "You are about to add the following users to group:`r`n$userList`r`nGroup ID: $groupObjectId`r`nTAP validity: $tapHours hours`r`nProceed?"
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Add to Group", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        # --- Progress form for live status ---
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Processing TAP and Group Add"
        $progressForm.Size = New-Object System.Drawing.Size(600,400)
        $txtProgress = New-Object System.Windows.Forms.TextBox
        $txtProgress.Multiline = $true
        $txtProgress.ScrollBars = "Vertical"
        $txtProgress.ReadOnly = $true
        $txtProgress.Dock = "Fill"
        $progressForm.Controls.Add($txtProgress)
        $progressForm.Show()

        foreach ($sel in $lstUsers.SelectedItems) {
            $userEmail = $sel.Text
            $dt = Get-Date -Format "yyyy-MM-dd HH:mm"
            $batchName = $txtBatchName.Text
            $txtProgress.AppendText("Processing $userEmail...`r`n")
            # Get user object ID and names
            try {
                $txtProgress.AppendText("Fetching user object...`r`n")
                $userObj = Get-MgUser -UserId $userEmail
                $userId = $userObj.Id
                $firstName = $userObj.GivenName
                $lastName = $userObj.Surname
                if (-not $firstName -or -not $lastName) {
                    $userPart = ($userEmail -split "@")[0]
                    if ($userPart -match "^(.*?)[\._\- ](.*)$") {
                        $firstName = $matches[1]
                        $lastName = $matches[2]
                    } else {
                        $firstName = $userPart
                        $lastName = ""
                    }
                }
            } catch {
                $errMsg = $_.Exception.Message
                $txtProgress.AppendText("Could not find user: ${userEmail}: $errMsg`r`n")
                continue
            }
            # Resolve user objectId from UPN/email
            $userObj = Get-MgUser -Filter "userPrincipalName eq '$userEmail'"
            if (!$userObj) {
                $txtProgress.AppendText("User $userEmail not found in Entra/Azure AD.`r`n")
                $sel.SubItems[3].Text = "Not found"
                $sel.SubItems[3].ForeColor = [System.Drawing.Color]::Red
                Add-Content -Path $logFile -Value "$dt,${batchName},${userEmail},User not found"
                continue
            }
            $userId = $userObj.Id

            # Add user to group
            $addedToGroup = $false
            try {
                $txtProgress.AppendText("Adding $userEmail to group $groupObjectId...`r`n")
                New-MgGroupMember -GroupId $groupObjectId -DirectoryObjectId $userId -ErrorAction Stop
                $addedToGroup = $true
                $txtProgress.AppendText("Added to group.`r`n")
            } catch {
                $errMsg = $_.Exception.Message
                if ($errMsg -notmatch "added object references already exist" -and $errMsg -notmatch "One or more added object references already exist") {
                    $txtProgress.AppendText("Error adding ${userEmail} to group: $errMsg`r`n")
                } else {
                    $addedToGroup = $true
                    $txtProgress.AppendText("Already a member of group.`r`n")
                }
            }
            if ($addedToGroup) {
                Add-Content -Path $logFile -Value "$dt,${batchName},${userEmail},Added to Group"
                $sel.SubItems[3].Text = "Yes"
                $sel.SubItems[3].ForeColor = [System.Drawing.Color]::DarkGreen
            }

            # Create TAP
            $tapCreated = $false
            try {
                $txtProgress.AppendText("Generating TAP for $tapHours hours...`r`n")
                $tapMethod = New-MgUserAuthenticationTemporaryAccessPassMethod `
                    -UserId $userId `
                    -BodyParameter @{lifetimeInMinutes=$maxTAP; isUsableOnce=$true}
                $tap = $tapMethod.TemporaryAccessPass
                Add-Content -Path $logFile -Value "$dt,${batchName},${userEmail},TAP Set"
                $sel.SubItems[2].Text = "Yes"
                $sel.SubItems[2].ForeColor = [System.Drawing.Color]::DodgerBlue
                $tapCreated = $true
                $txtProgress.AppendText("TAP created.`r`n")
            } catch {
                $errMsg = $_.Exception.Message
                $txtProgress.AppendText("Error creating TAP for ${userEmail}: $errMsg`r`n")
                continue
            }
            # Prepare HTML email
            $templatePath = "TAP_Email_Template.html"
            if (-not (Test-Path $templatePath)) {
                $txtProgress.AppendText("Email template not found: ${templatePath}`r`n")
                continue
            }
            $html = Get-Content $templatePath -Raw
            $html = $html -replace "\{FirstName\}", $firstName
            $html = $html -replace "\{LastName\}", $lastName
            $html = $html -replace "\{TAP\}", $tap
            $html = $html -replace "\{TAP_HOURS\}", $tapHours

            # Send email
            if ($tapCreated) {
                try {
                    $txtProgress.AppendText("Emailing TAP to $userEmail...`r`n")
                    Send-MgUserMail -UserId $userId -BodyParameter @{
                        message = @{
                            subject = "Your Temporary Access Pass (TAP) for Passkey Setup"
                            body = @{
                                contentType = "html"
                                content = $html
                            }
                            toRecipients = @(@{emailAddress = @{address = $userEmail}})
                        }
                        saveToSentItems = $false
                    }
                    Add-Content -Path $logFile -Value "$dt,${batchName},${userEmail},TAP sent"
                    $sel.SubItems[4].Text = "Yes"
                    $sel.SubItems[4].ForeColor = [System.Drawing.Color]::Blue
                    $txtProgress.AppendText("Email sent to $userEmail.`r`n")
                } catch {
                    $errMsg = $_.Exception.Message
                    $txtProgress.AppendText("Could not send email to ${userEmail}: $errMsg`r`n")
                    continue
                }
            }
            $txtProgress.AppendText("Done for $userEmail.`r`n`r`n")
        }
        $txtProgress.AppendText("All done!`r`n")
        $lblSaved.Text = "TAP sent for selected users and recorded."
    })

    $batchForm.Add_Shown({
        if ($defaults["last_batch"] -and -not $batchToLoad) {
            $resumeResult = [System.Windows.Forms.MessageBox]::Show(
                "Would you like to resume your last batch '$($defaults["last_batch"])?'",
                "Resume Batch",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($resumeResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                $txtBatchName.Text = $defaults["last_batch"]
                Refresh-ListView $defaults["last_batch"]
            } else {
                $txtBatchName.Text = ""
                $lstUsers.Items.Clear()
            }
        }
    })

    [void]$batchForm.ShowDialog()
}

# --- Main Connect Logic and Script Launcher ---
$btnConnect.Add_Click({
    $txtStatus.Text = ""
    $tenantName = $txtTenantName.Text.Trim()
    $appId     = $txtAppId.Text.Trim()
    $tenantId  = $txtTenantId.Text.Trim()
    $tapGroup  = $txtTapGroup.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($tenantName) -or [string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($tenantId)) {
        $txtStatus.Text = "Tenant Name, App ID, and Tenant ID are required."
        return
    }
    $defaults["tap_group_objectid"] = $tapGroup
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

[void]$form.ShowDialog()

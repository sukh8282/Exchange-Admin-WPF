<#  Exchange-Admin-WPF.ps1  (Compat + Pro UI)
    - WPF UI for Exchange admin actions (WinPS 5.1 / STA)
    - Data-driven Actions registry (add actions with a single entry)
    - All outputs -> sortable DataGrid + filter box + context menu
    - Menu bar: File / View / Help + Connect/Disconnect buttons
    - Safe async harness available (default runs sync to avoid session exits)
    - Indicators: EXO connection + ImportExcel presence
    - Settings persisted to %APPDATA%\ExAdminWpf\settings.json
#>

# ===== Environment checks =====
Add-Type -AssemblyName System.Windows.Forms
if ($PSVersionTable.PSEdition -ne 'Desktop') {
  [System.Windows.Forms.MessageBox]::Show(
    "Run in Windows PowerShell 5.1 (Desktop), not PowerShell 7 (pwsh).",
    "Exchange Admin WPF", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
  return
}
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  [System.Windows.Forms.MessageBox]::Show(
    "Run in STA mode. Example:`n`npowershell.exe -NoExit -STA -ExecutionPolicy Bypass -File .\Exchange-Admin-WPF.ps1",
    "Exchange Admin WPF", [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
  return
}

# ===== WPF assemblies =====
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase

# ===== Globals / helpers =====
$ErrorActionPreference = 'Continue'
$AppName      = 'ExAdminWpf'
$settingsPath = Join-Path $env:APPDATA "$AppName\settings.json"

function Ensure-Dir([string]$Path){ if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null } }
Ensure-Dir "C:\Temp"

function Ensure-ExcelModule {
  if (-not (Get-Module -ListAvailable -Name ImportExcel)) { return $false }
  if (-not (Get-Module -Name ImportExcel)) { Import-Module ImportExcel -ErrorAction SilentlyContinue | Out-Null }
  return $true
}
function Test-ExchangeOnline {
  try { Get-OrganizationConfig -ErrorAction Stop | Out-Null; return $true } catch { return $false }
}
function Ensure-EXOModule {
  if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) { return $false }
  if (-not (Get-Module -Name ExchangeOnlineManagement)) { Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue | Out-Null }
  return $true
}
function Update-ConnStatus {
  if ($lblConn) {
    if (Test-ExchangeOnline) { $lblConn.Text="Connected"; $dotConn.Fill=[Windows.Media.Brushes]::LimeGreen }
    else { $lblConn.Text="Not connected"; $dotConn.Fill=[Windows.Media.Brushes]::Tomato }
  }
}
function Connect-EXO {
  if (-not (Ensure-EXOModule)) { [System.Windows.Forms.MessageBox]::Show("ExchangeOnlineManagement module not available","Connect") | Out-Null; return }
  try { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Connect") | Out-Null } finally { Update-ConnStatus }
}
function Disconnect-EXO { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} ; Update-ConnStatus }

# Present data in grid (wrap strings/scalars)
function Present-Data($data) {
  if ($null -eq $data)          { $gridOut.ItemsSource = ,([pscustomobject]@{ Info="No output." }); return }
  if ($data -is [string])       { $gridOut.ItemsSource = ,([pscustomobject]@{ Info=$data }); return }
  if ($data -is [ValueType])    { $gridOut.ItemsSource = ,([pscustomobject]@{ Value="$data" }); return }
  if ($data -is [System.Collections.IEnumerable] -and $data -isnot [string]) {
    $arr = @(); foreach ($i in $data) { if ($i -is [string] -or $i -is [ValueType]) { $arr += [pscustomobject]@{ Value="$i" } } else { $arr += $i } }
    $gridOut.ItemsSource = $arr; return
  }
  $gridOut.ItemsSource = ,$data
}
function Set-Busy($on){ if ($btnRun) { $btnRun.IsEnabled = -not $on }; if ($prg) { $prg.Visibility = if($on){'Visible'} else {'Collapsed'} }; if ($lblStatus) { $lblStatus.Text = if($on){'Running…'} else {'Ready'} } }
function Set-Vis($panel, [bool]$show) { if ($null -eq $panel) { return }; $panel.Visibility = if ($show) { 'Visible' } else { 'Collapsed' } }
# ===== Async engine (WinPS 5.1 safe; optional) =====
Add-Type -AssemblyName System.Core
$runspacePool = [runspacefactory]::CreateRunspacePool(1,4)
$runspacePool.Open()
function Invoke-Async([scriptblock]$Script, [hashtable]$Params){
  $ps = [powershell]::Create().AddScript($Script).AddParameters($Params)
  $ps.RunspacePool = $runspacePool
  Set-Busy $true
  $handle = $ps.BeginInvoke()
  [void][System.Threading.ThreadPool]::RegisterWaitForSingleObject(
    $handle.AsyncWaitHandle,
    {
      param($state,$timedOut)
      $psLocal = $state.PS; $hLocal  = $state.Handle
      try { $result = $psLocal.EndInvoke($hLocal) } catch { $result = ,([pscustomobject]@{ Error = $_.Exception.Message }) } finally { $psLocal.Dispose() }
      try { $state.Grid.Dispatcher.Invoke({ Present-Data $args[0]; Set-Busy $false }, @($result)) } catch {}
    },
    @{ PS=$ps; Handle=$handle; Grid=$gridOut },
    -1,$true)
}

# ===== Settings =====
function Load-Settings {
  if (Test-Path $settingsPath) { try { return Get-Content $settingsPath -Raw | ConvertFrom-Json } catch { } }
  return [pscustomobject]@{ MenuIndex=0; Primary=""; Secondary=""; UseAsyncForHeavy=$false; Window=@{} }
}
function Save-Settings {
  try {
    $dir = Split-Path $settingsPath
    if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory | Out-Null }
    $obj = [pscustomobject]@{
      MenuIndex = $menuList.SelectedIndex
      Primary   = $txtPrimary.Text
      Secondary = $txtSecondary.Text
      UseAsyncForHeavy = $script:Settings.UseAsyncForHeavy
      Window    = @{ Width=$win.Width; Height=$win.Height; Left=$win.Left; Top=$win.Top }
    }
    $obj | ConvertTo-Json | Set-Content $settingsPath
  } catch { }
}
$script:Settings = Load-Settings

# ===== Shared helpers =====
function Convert-SizeStringToMB([string]$s) {
  if (-not $s) { return 0 }
  if ($s -match '([0-9.]+)\s*(KB|MB|GB|TB)') {
    $v = [double]$matches[1]
    switch ($matches[2].ToUpper()) {
      'KB' { return [math]::Round($v / 1024, 2) }
      'MB' { return [math]::Round($v, 2) }
      'GB' { return [math]::Round($v * 1024, 2) }
      'TB' { return [math]::Round($v * 1024 * 1024, 2) }
    }
  }
  return 0
}
# ===== Action implementations (Act-*) =====
function Act-1([string]$Mailbox){
  $fa = Get-MailboxPermission -Identity $Mailbox -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains "FullAccess" } |
        Select-Object @{N='Type';E={'FullAccess'}}, @{N='User';E={$_.User}}, @{N='Rights';E={$_.AccessRights -join ', '}}
  $sa = Get-RecipientPermission -Identity $Mailbox -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains "SendAs" } |
        Select-Object @{N='Type';E={'SendAs'}},   @{N='User';E={$_.Trustee}}, @{N='Rights';E={$_.AccessRights -join ', '}}
  $sob = (Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue).GrantSendOnBehalfTo |
        ForEach-Object { [pscustomobject]@{ Type='SendOnBehalf'; User = $_; Rights='GrantSendOnBehalfTo' } }
  if ($fa -or $sa -or $sob) { $fa + $sa + $sob } else { [pscustomobject]@{ Info="No permissions found"; Mailbox=$Mailbox } }
}
function Act-2([string]$Mailbox,[string]$User,[string]$Kind){
  switch($Kind){
    'FullAccess'   { Add-MailboxPermission -Identity $Mailbox -User $User -AccessRights FullAccess }
    'SendAs'       { Add-RecipientPermission -Identity $Mailbox -Trustee $User -AccessRights SendAs }
    'SendOnBehalf' { Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{Add=$User} }
  }
  [pscustomobject]@{ Action='AddPermission'; Mailbox=$Mailbox; User=$User; Kind=$Kind; Status='OK' }
}
function Act-3([string]$Mailbox,[string]$User,[string]$Kind){
  switch($Kind){
    'FullAccess'   { Remove-MailboxPermission -Identity $Mailbox -User $User -AccessRights FullAccess -Confirm:$false }
    'SendAs'       { Remove-RecipientPermission -Identity $Mailbox -Trustee $User -AccessRights SendAs -Confirm:$false }
    'SendOnBehalf' { Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{Remove=$User} }
  }
  [pscustomobject]@{ Action='RemovePermission'; Mailbox=$Mailbox; User=$User; Kind=$Kind; Status='OK' }
}
# Calendar
function Cal-View([string]$Mailbox){
  Get-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -ErrorAction SilentlyContinue |
    Select-Object @{N='User';E={$_.User}}, @{N='AccessRights';E={$_.AccessRights -join ', '}}, @{N='Flags';E={$_.SharingPermissionFlags -join ', '}}
}
function Cal-Remove([string]$Mailbox,[string]$User){
  Remove-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -User $User -Confirm:$false
  [pscustomobject]@{ Action='Cal-Remove'; Mailbox=$Mailbox; User=$User; Status='OK' }
}
function Cal-Add([string]$Mailbox,[string]$User,[string]$Rights){
  Add-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -User $User -AccessRights $Rights
  [pscustomobject]@{ Action='Cal-Add'; Mailbox=$Mailbox; User=$User; Rights=$Rights; Status='OK' }
}
function Cal-ResetDelegates([string]$Mailbox){
  Remove-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -ResetDelegateUserCollection -Confirm:$false
  [pscustomobject]@{ Action='Cal-ResetDelegates'; Mailbox=$Mailbox; Status='OK' }
}
function Cal-SetDelegate([string]$Mailbox,[string]$User,[string]$Rights){
  Set-MailboxFolderPermission -Identity "${Mailbox}:\Calendar" -User $User -AccessRights $Rights -SharingPermissionFlags Delegate -SendNotificationToUser:$false
  [pscustomobject]@{ Action='Cal-SetDelegate'; Mailbox=$Mailbox; User=$User; Rights=$Rights; Status='OK' }
}
# Stats / Retention / MFA
function Act-5-Stats([string]$Mailbox){
  Get-MailboxFolderStatistics -Identity $Mailbox |
    Select-Object Identity, ItemsInFolder, FolderSize, FolderPath, FolderType
}
function Act-5-ApplyRetention([string]$Mailbox){
  $mbx = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
  $current = $mbx.RetentionPolicy
  $pols = Get-RetentionPolicy | Select -ExpandProperty Name
  Add-Type -AssemblyName Microsoft.VisualBasic
  $choice = [Microsoft.VisualBasic.Interaction]::InputBox("Enter EXACT policy name to apply:`n`n$($pols -join "`n")",
            "Retention Policy - current: $current", ($pols | Select-Object -First 1))
  if ($choice) {
    Set-Mailbox -Identity $Mailbox -RetentionPolicy $choice
    [pscustomobject]@{ Action='ApplyRetention'; Mailbox=$Mailbox; Policy=$choice; Status='OK' }
  } else {
    [pscustomobject]@{ Action='ApplyRetention'; Mailbox=$Mailbox; Policy=$current; Status='Cancelled' }
  }
}
function Act-5-RunMFA([string]$Mailbox){
  Start-ManagedFolderAssistant -Identity $Mailbox
  [pscustomobject]@{ Action='RunMFA'; Mailbox=$Mailbox; Timestamp=(Get-Date); Status='Triggered' }
}
# DG manage
function DG-View([string]$DG){
  $d = Get-DistributionGroup -Identity $DG -ErrorAction Stop
  $owners = $d.ManagedBy | ForEach-Object { Get-Recipient $_ | Select-Object DisplayName,PrimarySmtpAddress }
  $members = Get-DistributionGroupMember -Identity $DG | Select-Object DisplayName,PrimarySmtpAddress
  $summary = [pscustomobject]@{ Type='Group'; DisplayName=$d.DisplayName; Description=$d.Notes }
  $summary, ($owners | ForEach-Object { [pscustomobject]@{ Type='Owner'; DisplayName=$_.DisplayName; Email=$_.PrimarySmtpAddress } }),
           ($members| ForEach-Object { [pscustomobject]@{ Type='Member';DisplayName=$_.DisplayName; Email=$_.PrimarySmtpAddress } })
}
function DG-AddMember([string]$DG,[string]$User){
  Add-DistributionGroupMember -Identity $DG -Member $User -ErrorAction Stop
  [pscustomobject]@{ Action='DG-AddMember'; Group=$DG; Member=$User; Status='OK' }
}
function DG-RemoveMember([string]$DG,[string]$User){
  Remove-DistributionGroupMember -Identity $DG -Member $User -Confirm:$false -ErrorAction Stop
  [pscustomobject]@{ Action='DG-RemoveMember'; Group=$DG; Member=$User; Status='OK' }
}
# OOF
function OOF-Check([string]$Mailbox){
  $o = Get-MailboxAutoReplyConfiguration -Identity $Mailbox -ErrorAction SilentlyContinue
  [pscustomobject]@{
    Mailbox=$Mailbox; AutoReplyState=$o.AutoReplyState; StartTime=$o.StartTime; EndTime=$o.EndTime;
    ExternalAudience=$o.ExternalAudience; InternalMessage=$o.InternalMessage; ExternalMessage=$o.ExternalMessage
  }
}
function OOF-SetScheduled([string]$Mailbox,[datetime]$Start,[datetime]$End,[string]$Int,[string]$Ext){
  Set-MailboxAutoReplyConfiguration -Identity $Mailbox -AutoReplyState Scheduled -StartTime $Start -EndTime $End -InternalMessage $Int -ExternalMessage $Ext -ExternalAudience All
  [pscustomobject]@{ Action='OOF-SetScheduled'; Mailbox=$Mailbox; Start=$Start; End=$End; Status='OK' }
}
function OOF-SetIndef([string]$Mailbox,[string]$Int,[string]$Ext){
  Set-MailboxAutoReplyConfiguration -Identity $Mailbox -AutoReplyState Enabled -InternalMessage $Int -ExternalMessage $Ext -ExternalAudience All
  [pscustomobject]@{ Action='OOF-SetIndefinite'; Mailbox=$Mailbox; Status='OK' }
}
function OOF-Off([string]$Mailbox){
  Set-MailboxAutoReplyConfiguration -Identity $Mailbox -AutoReplyState Disabled
  [pscustomobject]@{ Action='OOF-Off'; Mailbox=$Mailbox; Status='OK' }
}
# Option 8: Mailbox Stats DETAIL
function Act-8([string]$Mailbox){
  Ensure-Dir "C:\Temp"; $ok = Ensure-ExcelModule
  $date = Get-Date -Format "yyyy-MM-dd_HHmmss"
  $outFile = "C:\Temp\$Mailbox-MailboxStats-$date.xlsx"
  $raw = Get-MailboxFolderStatistics -Identity $Mailbox |
         Select Name, Identity, ItemsInFolder, FolderSize, FolderAndSubfolderSize, FolderPath
  $tidy = $raw | ForEach-Object {
    $folderSizeDisplay  = ($_.FolderSize -replace '\s*\(.+\)$','')
    $totalSizeDisplay   = ($_.FolderAndSubfolderSize -replace '\s*\(.+\)$','')
    [pscustomobject]@{
      Name        = $_.Name
      Items       = $_.ItemsInFolder
      FolderSize  = $folderSizeDisplay
      TotalSize   = $totalSizeDisplay
      TotalMB     = Convert-SizeStringToMB $totalSizeDisplay
      Path        = $_.FolderPath
    }
  } | Sort-Object -Property TotalMB -Descending
  if ($ok -and $tidy.Count -gt 0) {
    try { $tidy | Select Name, Items, FolderSize, TotalSize, Path | Export-Excel -Path $outFile -WorksheetName "Mailbox Stats" -AutoSize -Title "Mailbox Folder Statistics" } catch {}
  }
  $tidy
}
# Option 9: DL + nested
function Act-9([string]$DL){
  $ok = Ensure-ExcelModule
  $OutputPath = "C:\temp\DL_Members_$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"
  function Get-MembersRecursively([string]$DLIdentity){
    try { $DLMembers = Get-DistributionGroupMember -Identity $DLIdentity -ResultSize Unlimited } catch { return @() }
    $acc=@()
    foreach($m in $DLMembers){
      $SyncType = if ($m.IsDirSynced) { "On-Prem Sync" } else { "Cloud" }
      $acc += [pscustomobject]@{
        ParentGroup=$DLIdentity; MemberType=$m.RecipientType; DisplayName=$m.DisplayName; Email=$m.PrimarySmtpAddress; SyncType=$SyncType
      }
      if ($m.RecipientType -in 'MailUniversalSecurityGroup','MailUniversalDistributionGroup'){
        $acc += Get-MembersRecursively -DLIdentity $m.PrimarySmtpAddress
      }
    }
    $acc
  }
  $all = Get-MembersRecursively -DLIdentity $DL
  if ($ok -and $all.Count -gt 0) { try { $all | Export-Excel -Path $OutputPath -AutoSize } catch {} }
  $all
}
# Option 10: Find Recipient
function Act-10([string]$Email){
  $r = Get-Recipient -RecipientPreviewFilter "EmailAddresses -eq 'SMTP:$Email'"
  if ($r) { $r | Select Name,PrimarySmtpAddress,RecipientType,RecipientTypeDetails } else { [pscustomobject]@{ Info="No recipient found"; Email=$Email } }
}
# Option 11: Inbox rules
function Act-11([string]$Mailbox){
  $ok = Ensure-ExcelModule
  $OutputFile = "C:\Temp\InboxRules_$($Mailbox)_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
  $rules = Get-InboxRule -Mailbox $Mailbox
  $processed=@()
  foreach($r in $rules){
    foreach($p in ($r | Get-Member -MemberType Properties)){
      $name=$p.Name; $val=$r.$name
      if ($null -ne $val -and $val -ne "" -and $val -notlike "{}" -and $val -ne $false -and $val -ne 0){
        $processed += [pscustomobject]@{ Rule=$r.Name; Attribute=$name; Value=($val -join ", ") }
      }
    }
  }
  if ($ok -and $processed.Count -gt 0) { try { $processed | Export-Excel -Path $OutputFile -AutoSize -Title "Inbox Rules for $Mailbox" -WorksheetName "Rules" } catch {} }
  $processed
}
# Option 12: PF permissions
function Act-12([string]$Path){
  $Permissions = Get-PublicFolderClientPermission -Identity $Path
  if (!$Permissions){ return [pscustomobject]@{ Info="No permissions"; Path=$Path } }
  $rows=@()
  function Process-GroupMembers([string]$GroupName,[string]$Parent,[string]$Perm){
    $Members = Get-DistributionGroupMember -Identity $GroupName | Select DisplayName,PrimarySmtpAddress,RecipientTypeDetails
    $acc=@()
    foreach ($m in $Members) {
      if ($m.RecipientTypeDetails -like "*Group") {
        $acc += [pscustomobject]@{ Type="NestedGroup"; ParentGroup=$Parent; GroupName=$GroupName; Name=$m.DisplayName; Email=$m.PrimarySmtpAddress; Permission=$Perm }
        $acc += Process-GroupMembers -GroupName $m.DisplayName -Parent $GroupName -Perm $Perm
      } else {
        $acc += [pscustomobject]@{ Type="Member"; ParentGroup=$Parent; GroupName=$GroupName; Name=$m.DisplayName; Email=$m.PrimarySmtpAddress; Permission=$Perm }
      }
    }
    $acc
  }
  foreach ($perm in $Permissions) {
    try {
      $rec = Get-Recipient -Identity $perm.User -ErrorAction Stop
      if ($rec.RecipientTypeDetails -like "*Group") {
        $rows += Process-GroupMembers -GroupName $rec.Name -Parent "" -Perm ($perm.AccessRights -join ", ")
      } else {
        $rows += [pscustomobject]@{ Type="User"; ParentGroup=""; GroupName=""; Name=$perm.User; Email=$rec.PrimarySmtpAddress; Permission=$perm.AccessRights -join ", " }
      }
    } catch {
      $rows += [pscustomobject]@{ Type="User"; ParentGroup=""; GroupName=""; Name=$perm.User; Email=""; Permission=$perm.AccessRights -join ", " }
    }
  }
  $rows
}
# Option 13: Quota
function Act-13([string]$Mailbox){
  Add-Type -AssemblyName Microsoft.VisualBasic
  function ParseQuota([string]$q){
    if ($q -eq "Unlimited"){ return "Unlimited" }
    if ($q -match "^(\d+(\.\d+)?)(\s*B|\s*KB|\s*MB|\s*GB|\s*TB)?$"){
      $v=[double]$matches[1]; $u=$matches[3]; if ([string]::IsNullOrWhiteSpace($u)) { $u="B" }; $u=$u.Trim().ToUpper()
      switch($u){ "B"{$v=$v/1GB};"KB"{$v=$v/1MB};"MB"{$v=$v/1GB};"GB"{$v=$v};"TB"{$v=$v*1024}; default{ throw "Invalid unit: $q" } }
      return "$v GB"
    } else { throw "Invalid quota value: $q. Use number + unit (B/KB/MB/GB/TB) or 'Unlimited'." }
  }
  $mbx = Get-Mailbox -Identity $Mailbox | Select DisplayName,ProhibitSendQuota,ProhibitSendReceiveQuota,IssueWarningQuota
  $iw = [Microsoft.VisualBasic.Interaction]::InputBox("IssueWarningQuota (e.g. 50 GB / Unlimited / blank=skip)","Quota",$mbx.IssueWarningQuota)
  $ps = [Microsoft.VisualBasic.Interaction]::InputBox("ProhibitSendQuota (e.g. 100 GB / Unlimited / blank=skip)","Quota",$mbx.ProhibitSendQuota)
  $pr = [Microsoft.VisualBasic.Interaction]::InputBox("ProhibitSendReceiveQuota (e.g. 150 GB / Unlimited / blank=skip)","Quota",$mbx.ProhibitSendReceiveQuota)
  $iw2 = if($iw){ ParseQuota $iw } else { $mbx.IssueWarningQuota }
  $ps2 = if($ps){ ParseQuota $ps } else { $mbx.ProhibitSendQuota }
  $pr2 = if($pr){ ParseQuota $pr } else { $mbx.ProhibitSendReceiveQuota }
  Set-Mailbox -Identity $Mailbox -IssueWarningQuota $iw2 -ProhibitSendQuota $ps2 -ProhibitSendReceiveQuota $pr2
  [pscustomobject]@{ Action='UpdateQuota'; Mailbox=$Mailbox; IssueWarning=$iw2; ProhibitSend=$ps2; ProhibitSendReceive=$pr2; Status='OK' }
}
# Option 14: DL Send-As
function Act-15([string]$DL){
  $p = Get-RecipientPermission -Identity $DL | Where-Object { $_.AccessRights -contains "SendAs" } |
       Select-Object @{N='Trustee';E={$_.Trustee}}, @{N='Rights';E={$_.AccessRights -join ', '}}
  if ($p) { $p } else { [pscustomobject]@{ Info="No Send-As permissions"; Group=$DL } }
}
# Option 15: DL Allowed Senders
function Act-16([string]$DL){
  function Resolve-Ids($ids){ $out=@(); foreach($id in $ids){ try{ $r=Get-Recipient -Identity $id -ErrorAction Stop; $out += $r.PrimarySmtpAddress } catch { $out += "Unknown (ID: $id)" } }; $out }
  $dl = Get-DistributionGroup -Identity $DL -ErrorAction Stop
  $rows=@()
  $map = @(
    @{Name='AcceptMessagesOnlyFrom'; Val=$dl.AcceptMessagesOnlyFrom},
    @{Name='AcceptMessagesOnlyFromWithDisplayNames'; Val=$dl.AcceptMessagesOnlyFromWithDisplayNames},
    @{Name='AcceptMessagesOnlyFromSendersOrMembers'; Val=$dl.AcceptMessagesOnlyFromSendersOrMembers},
    @{Name='AcceptMessagesOnlyFromSendersOrMembersWithDisplayNames'; Val=$dl.AcceptMessagesOnlyFromSendersOrMembersWithDisplayNames},
    @{Name='AcceptMessagesOnlyFromDLMembers'; Val=$dl.AcceptMessagesOnlyFromDLMembers},
    @{Name='AcceptMessagesOnlyFromDLMembersWithDisplayNames'; Val=$dl.AcceptMessagesOnlyFromDLMembersWithDisplayNames}
  )
  foreach($entry in $map){
    $vals = $entry.Val
    if ($vals -and $vals.Count -gt 0){
      $emails = Resolve-Ids $vals
      foreach($e in $emails){ $rows += [pscustomobject]@{ Setting=$entry.Name; Email=$e } }
    }
  }
  if ($rows.Count -gt 0){ $rows } else { [pscustomobject]@{ Info="No restrictions (open to all senders)"; Group=$DL } }
}
# Option 16: Remove user from DG
function Act-17([string]$DG,[string]$User){
  Remove-DistributionGroupMember -Identity $DG -Member $User -Confirm:$false
  [pscustomobject]@{ Action='RemoveFromDG'; Group=$DG; User=$User; Status='OK' }
}
function Act-18-InactiveMailboxes {
  try {
    $inactive = Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction Stop
    $inactive | Select-Object DisplayName, PrimarySmtpAddress, WhenSoftDeleted, RecipientTypeDetails
  } catch {
    ,([pscustomobject]@{ Error = $_.Exception.Message })
  }
}

# ===== Actions registry (data-driven) =====
$Actions = @(
  @{ Key=1;  Menu="1. View mailbox permissions"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Mailbox → FullAccess / Send-As / Send-On-Behalf" }; Heavy=$false; Run={ param($ctx) Act-1 -Mailbox $ctx.Primary } },
  @{ Key=2;  Menu="2. Add permissions to a mailbox"; UI=@{ Primary=$true; Secondary=$true; Option=$true; Options=@("FullAccess","SendOnBehalf","SendAs"); Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Add permission. Choose type." }; Heavy=$false; Run={ param($ctx) if (-not $ctx.Secondary){ [pscustomobject]@{ Info='Enter User' } } else { Act-2 -Mailbox $ctx.Primary -User $ctx.Secondary -Kind $ctx.Option } } },
  @{ Key=3;  Menu="3. Remove permissions from a mailbox"; UI=@{ Primary=$true; Secondary=$true; Option=$true; Options=@("FullAccess","SendOnBehalf","SendAs"); Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Remove permission. Choose type." }; Heavy=$false; Run={ param($ctx) if (-not $ctx.Secondary){ [pscustomobject]@{ Info='Enter User' } } else { Act-3 -Mailbox $ctx.Primary -User $ctx.Secondary -Kind $ctx.Option } } },
  @{ Key=4;  Menu="4. View calendar permissions"; UI=@{ Primary=$true; Secondary=$true; Option=$true; Options=@("View","Remove","Add","ResetDelegates","SetDelegate"); Extra=$true; StartEnd=$false; OOFMsg=$false; Hint="Calendar: View/Remove/Add/ResetDelegates/SetDelegate" }; Heavy=$false; Run={ param($ctx)
        switch($ctx.Option){
          'View'            { Cal-View -Mailbox $ctx.Primary }
          'Remove'          { if(-not $ctx.Secondary){ [pscustomobject]@{ Info='Enter User' } } else { Cal-Remove -Mailbox $ctx.Primary -User $ctx.Secondary } }
          'Add'             { if(-not $ctx.Secondary -or -not $ctx.Extra){ [pscustomobject]@{ Info='Enter User and AccessRights' } } else { Cal-Add -Mailbox $ctx.Primary -User $ctx.Secondary -Rights $ctx.Extra } }
          'ResetDelegates'  { Cal-ResetDelegates -Mailbox $ctx.Primary }
          'SetDelegate'     { if(-not $ctx.Secondary -or -not $ctx.Extra){ [pscustomobject]@{ Info='Enter User and AccessRights' } } else { Cal-SetDelegate -Mailbox $ctx.Primary -User $ctx.Secondary -Rights $ctx.Extra } }
        } } },
  @{ Key=5;  Menu="5. Run mailbox statistics or retention policy or MFA"; UI=@{ Primary=$true; Secondary=$false; Option=$true; Options=@("Statistics","ApplyRetention","RunMFA"); Extra=$true; StartEnd=$false; OOFMsg=$false; Hint="Mailbox ops: Statistics / ApplyRetention / RunMFA (one-shot)"}; Heavy=$false; Run={ param($ctx)
        switch($ctx.Option){
          'Statistics'     { Act-5-Stats -Mailbox $ctx.Primary }
          'ApplyRetention' { Act-5-ApplyRetention -Mailbox $ctx.Primary }
          'RunMFA'         { Act-5-RunMFA -Mailbox $ctx.Primary }
        } } },
  @{ Key=6;  Menu="6. Manage Distribution Group"; UI=@{ Primary=$true; Secondary=$true; Option=$true; Options=@("View","AddMember","RemoveMember"); Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="DG manage: View / AddMember / RemoveMember" }; Heavy=$false; Run={ param($ctx)
        switch($ctx.Option){
          'View'         { DG-View -DG $ctx.Primary }
          'AddMember'    { if(-not $ctx.Secondary){ [pscustomobject]@{ Info='Enter Member SMTP' } } else { DG-AddMember -DG $ctx.Primary -User $ctx.Secondary } }
          'RemoveMember' { if(-not $ctx.Secondary){ [pscustomobject]@{ Info='Enter Member SMTP' } } else { DG-RemoveMember -DG $ctx.Primary -User $ctx.Secondary } }
        } } },
  @{ Key=7;  Menu="7. OOF Config"; UI=@{ Primary=$true; Secondary=$false; Option=$true; Options=@("Check","SetScheduled","SetIndefinite","TurnOff"); Extra=$false; StartEnd=$true; OOFMsg=$true; Hint="OOF: Check / SetScheduled / SetIndefinite / TurnOff" }; Heavy=$false; Run={ param($ctx)
        switch($ctx.Option){
          'Check'         { OOF-Check -Mailbox $ctx.Primary }
          'SetScheduled'  { if(-not $ctx.End){ [pscustomobject]@{ Info='Enter End time' } } else { $st = if($ctx.Start){ [datetime]$ctx.Start } else { Get-Date }; $et=[datetime]$ctx.End; OOF-SetScheduled -Mailbox $ctx.Primary -Start $st -End $et -Int $ctx.MsgInt -Ext $ctx.MsgExt } }
          'SetIndefinite' { OOF-SetIndef -Mailbox $ctx.Primary -Int $ctx.MsgInt -Ext $ctx.MsgExt }
          'TurnOff'       { OOF-Off -Mailbox $ctx.Primary }
        } } },
  @{ Key=8;  Menu="8. Get Mailbox Stats DETAIL"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Mailbox Stats DETAIL → Grid + optional Excel" }; Heavy=$true; Run={ param($ctx) Act-8 -Mailbox $ctx.Primary } },
  @{ Key=9;  Menu="9. Get DL and Nested Members"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="DL & nested members → Grid + optional Excel" }; Heavy=$true; Run={ param($ctx) Act-9 -DL $ctx.Primary } },
  @{ Key=10; Menu="10. Find Recipient"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Find Recipient by SMTP" }; Heavy=$false; Run={ param($ctx) Act-10 -Email $ctx.Primary } },
  @{ Key=11; Menu="11. Get Mailbox Rules"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Inbox Rules → Grid + optional Excel" }; Heavy=$true; Run={ param($ctx) Act-11 -Mailbox $ctx.Primary } },
  @{ Key=12; Menu="12. Get PF Permissions"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Public Folder permissions; PF path like \Offices\..." }; Heavy=$true; Run={ param($ctx) Act-12 -Path $ctx.Primary } },
  @{ Key=13; Menu="13. Increase mxb Quota"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Increase mailbox quota (dialog prompts)"}; Heavy=$false; Run={ param($ctx) Act-13 -Mailbox $ctx.Primary } },
  @{ Key=14; Menu="14. Check Users Mail-Enabled Groups (skipped – no Graph)"; UI=@{ Primary=$false; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="(Skipped to avoid Graph)" }; Heavy=$false; Run={ param($ctx) [pscustomobject]@{ Info='Skipped (no Graph).' } } },
  @{ Key=15; Menu="15. Check DL Send-As Permissions"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="DL Send-As → Grid" }; Heavy=$false; Run={ param($ctx) Act-15 -DL $ctx.Primary } },
  @{ Key=16; Menu="16. Check DL Allowed Senders"; UI=@{ Primary=$true; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="DL Allowed Senders → Grid" }; Heavy=$false; Run={ param($ctx) Act-16 -DL $ctx.Primary } },
  @{ Key=17; Menu="17. Remove User from DG"; UI=@{ Primary=$true; Secondary=$true; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false; Hint="Remove user from DG" }; Heavy=$false; Run={ param($ctx) if(-not $ctx.Secondary){ [pscustomobject]@{ Info='Enter User' } } else { Act-17 -DG $ctx.Primary -User $ctx.Secondary } } },
  @{ Key=18; Menu="18. Inactive Mailboxes"; UI=@{ Hint="Lists all inactive mailboxes in the tenant."; Primary=$false; Secondary=$false; Option=$false; Extra=$false; StartEnd=$false; OOFMsg=$false }; Heavy=$true; Run={ param($ctx) Act-18-InactiveMailboxes } }
)


function Get-CurrentAction { if ($menuList.SelectedIndex -lt 0) { return $null } return $Actions[$menuList.SelectedIndex] }
function Refresh-Fields {
  $A = Get-CurrentAction; if (-not $A) { return }
  $p = $A.UI; $lblHint.Text = $p.Hint
  $cboOption.Items.Clear()
  if ($p.Option -and $p.Options) { $p.Options | ForEach-Object { [void]$cboOption.Items.Add($_) }; $cboOption.SelectedIndex = 0 } else { $cboOption.SelectedIndex = -1 }
  Set-Vis $pnlPrimary $p.Primary; Set-Vis $pnlSecondary $p.Secondary; Set-Vis $pnlOption $p.Option
  Set-Vis $pnlExtra $p.Extra; Set-Vis $pnlStart $p.StartEnd; Set-Vis $pnlEnd $p.StartEnd
  Set-Vis $pnlMsgInt $p.OOFMsg; Set-Vis $pnlMsgExt $p.OOFMsg
  if (-not $p.Secondary) { $txtSecondary.Text = "" }
  if (-not $p.Option)    { $cboOption.SelectedIndex = -1 }
  if (-not $p.Extra)     { $txtExtra.Text = "" }
  if (-not $p.StartEnd)  { $txtStart.Text = ""; $txtEnd.Text = "" }
  if (-not $p.OOFMsg)    { $txtMsgInt.Text = ""; $txtMsgExt.Text = "" }
}

# ===== XAML (compat-safe) =====
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Admin WPF" Height="780" Width="1250"
        WindowStartupLocation="CenterScreen" FontSize="13" Background="#FAFBFD">
  <DockPanel>
    <Menu DockPanel.Dock="Top">
      <MenuItem Header="_File">
        <MenuItem Header="_Connect to Exchange Online" x:Name="MnuConnect"/>
        <MenuItem Header="_Disconnect" x:Name="MnuDisconnect"/>
        <Separator/>
        <MenuItem Header="_Export Grid to CSV" x:Name="MnuExportCsv"/>
        <MenuItem Header="Export Grid to _Excel" x:Name="MnuExportXlsx"/>
        <Separator/>
        <MenuItem Header="E_xit" x:Name="MnuExit"/>
      </MenuItem>
      <MenuItem Header="_View">
        <MenuItem Header="_Refresh Fields" x:Name="MnuRefresh"/>
        <MenuItem Header="_Clear Output" x:Name="MnuClear"/>
        <MenuItem Header="Auto-size _Columns" x:Name="MnuAutosize"/>
      </MenuItem>
      <MenuItem Header="_Help">
        <MenuItem Header="_About" x:Name="MnuAbout"/>
        <MenuItem Header="Keyboard _Shortcuts" x:Name="MnuShortcuts"/>
      </MenuItem>
    </Menu>

    <Grid Margin="12,8,12,12">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="330"/>
        <ColumnDefinition Width="12"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <Border Grid.Column="0" BorderBrush="#DDE3EE" BorderThickness="1" CornerRadius="8" Padding="10" Background="White">
        <DockPanel>
          <StackPanel Orientation="Horizontal" DockPanel.Dock="Top" Margin="0,0,0,8">
            <TextBlock Text="Select Action" FontWeight="Bold" FontSize="14"/>
          </StackPanel>
          <ListBox x:Name="MenuList" DockPanel.Dock="Bottom" Height="680"/>
        </DockPanel>
      </Border>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="38"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <DockPanel Grid.Row="0" Margin="0,0,0,6">
          <TextBlock x:Name="LblHint" Foreground="#666" TextWrapping="Wrap" DockPanel.Dock="Left" Width="700"/>
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Ellipse x:Name="DotConn" Width="10" Height="10" Margin="0,0,6,0"/>
            <TextBlock x:Name="LblConn" Text="Not connected" Foreground="#666" Margin="0,0,12,0"/>
            <Ellipse x:Name="DotExcel" Width="10" Height="10" Margin="0,0,6,0"/>
            <TextBlock x:Name="LblExcel" Text="ImportExcel: missing" Foreground="#666" Margin="0,0,12,0"/>
            <Button x:Name="BtnConnect" Content="Connect" Width="90" Height="24" Margin="0,0,6,0"/>
            <Button x:Name="BtnDisconnect" Content="Disconnect" Width="90" Height="24"/>
          </StackPanel>
        </DockPanel>

        <StackPanel Grid.Row="1">
          <WrapPanel Margin="0,0,0,6">
            <StackPanel x:Name="PnlPrimary" Width="360" Margin="0,0,12,6">
              <TextBlock Text="Mailbox / DL / PF Path"/>
              <TextBox x:Name="TxtPrimary" ToolTip="Mailbox SMTP / DL address / Public Folder path (e.g. \Offices\...)" />
            </StackPanel>
            <StackPanel x:Name="PnlSecondary" Width="300" Margin="0,0,12,6">
              <TextBlock Text="User / Trustee / Member"/>
              <TextBox x:Name="TxtSecondary" ToolTip="User SMTP or UPN"/>
            </StackPanel>
            <StackPanel x:Name="PnlOption" Width="240" Margin="0,0,12,6">
              <TextBlock Text="Option"/>
              <ComboBox x:Name="CboOption"/>
            </StackPanel>
            <StackPanel x:Name="PnlExtra" Width="220" Margin="0,0,12,6">
              <TextBlock Text="Extra (Rights / Interval / Policy etc.)"/>
              <TextBox x:Name="TxtExtra"/>
            </StackPanel>
          </WrapPanel>
          <WrapPanel>
            <StackPanel x:Name="PnlStart" Width="240" Margin="0,0,12,6">
              <TextBlock Text="Start (YYYY-MM-DD HH:MM)"/>
              <TextBox x:Name="TxtStart"/>
            </StackPanel>
            <StackPanel x:Name="PnlEnd" Width="240" Margin="0,0,12,6">
              <TextBlock Text="End (YYYY-MM-DD HH:MM)"/>
              <TextBox x:Name="TxtEnd"/>
            </StackPanel>
            <StackPanel x:Name="PnlMsgInt" Width="340" Margin="0,0,12,6">
              <TextBlock Text="Internal Message (OOF)"/>
              <TextBox x:Name="TxtMsgInt"/>
            </StackPanel>
            <StackPanel x:Name="PnlMsgExt" Width="340" Margin="0,0,12,6">
              <TextBlock Text="External Message (OOF)"/>
              <TextBox x:Name="TxtMsgExt"/>
            </StackPanel>
          </WrapPanel>
        </StackPanel>

        <DockPanel Grid.Row="2">
          <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
            <Button x:Name="BtnRun"   Content="_Run" Width="140" Height="30" Margin="0,0,8,0"/>
            <Button x:Name="BtnClear" Content="C_lear" Width="100" Height="30" Margin="0,0,8,0"/>
            <Button x:Name="BtnCopy"  Content="_Copy Output" Width="160" Height="30"/>
          </StackPanel>
          <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" VerticalAlignment="Center">
            <TextBox x:Name="TxtFilter" Width="260" Height="28" Margin="0,0,12,0" ToolTip="Filter results"/>
            <TextBlock Text="Status:" FontWeight="Bold" Margin="0,0,6,0"/>
            <TextBlock x:Name="LblStatus" Text="Ready" Margin="0,0,12,0"/>
            <ProgressBar x:Name="Prg" Width="200" Height="10" Visibility="Collapsed" IsIndeterminate="True"/>
          </StackPanel>
        </DockPanel>

        <DataGrid x:Name="GridOut" Grid.Row="3"
                  AutoGenerateColumns="True"
                  CanUserSortColumns="True"
                  CanUserReorderColumns="True"
                  CanUserResizeColumns="True"
                  IsReadOnly="True"
                  AlternatingRowBackground="#FFF5F5F5"
                  HeadersVisibility="All"
                  GridLinesVisibility="Horizontal"
                  RowHeaderWidth="0" />
      </Grid>
    </Grid>
  </DockPanel>
</Window>
"@
# ===== Build window & controls =====
try { $win = [Windows.Markup.XamlReader]::Parse($xaml.OuterXml) } catch {
  [System.Windows.Forms.MessageBox]::Show("Failed to parse XAML:`r`n$($_.Exception.Message)","XAML error") | Out-Null
  return
}

$win.Title  = "Exchange Admin — $env:USERNAME@$env:USERDOMAIN"
$menuList   = $win.FindName("MenuList")
$txtPrimary = $win.FindName("TxtPrimary")
$txtSecondary= $win.FindName("TxtSecondary")
$cboOption  = $win.FindName("CboOption")
$txtExtra   = $win.FindName("TxtExtra")
$txtStart   = $win.FindName("TxtStart")
$txtEnd     = $win.FindName("TxtEnd")
$txtMsgInt  = $win.FindName("TxtMsgInt")
$txtMsgExt  = $win.FindName("TxtMsgExt")
$btnRun     = $win.FindName("BtnRun")
$btnClear   = $win.FindName("BtnClear")
$btnCopy    = $win.FindName("BtnCopy")
$txtFilter  = $win.FindName("TxtFilter")
$gridOut    = $win.FindName("GridOut")
$lblHint    = $win.FindName("LblHint")
$lblStatus  = $win.FindName("LblStatus")
$prg        = $win.FindName("Prg")
$dotConn    = $win.FindName("DotConn")
$lblConn    = $win.FindName("LblConn")
$dotExcel   = $win.FindName("DotExcel")
$lblExcel   = $win.FindName("LblExcel")
$btnConnect = $win.FindName("BtnConnect")
$btnDisconnect = $win.FindName("BtnDisconnect")
$mnuConnect     = $win.FindName("MnuConnect")
$mnuDisconnect  = $win.FindName("MnuDisconnect")
$mnuExit      = $win.FindName("MnuExit")
$mnuRefresh   = $win.FindName("MnuRefresh")
$mnuClear     = $win.FindName("MnuClear")
$mnuAutosize  = $win.FindName("MnuAutosize")
$mnuAbout     = $win.FindName("MnuAbout")
$mnuShortcuts = $win.FindName("MnuShortcuts")
$mnuExportCsv = $win.FindName("MnuExportCsv")
$mnuExportXlsx= $win.FindName("MnuExportXlsx")

# Panels
$pnlPrimary   = $win.FindName("PnlPrimary")
$pnlSecondary = $win.FindName("PnlSecondary")
$pnlOption    = $win.FindName("PnlOption")
$pnlExtra     = $win.FindName("PnlExtra")
$pnlStart     = $win.FindName("PnlStart")
$pnlEnd       = $win.FindName("PnlEnd")
$pnlMsgInt    = $win.FindName("PnlMsgInt")
$pnlMsgExt    = $win.FindName("PnlMsgExt")

# Init indicators
$dotConn.Fill  = [Windows.Media.Brushes]::Tomato
$dotExcel.Fill = [Windows.Media.Brushes]::Tomato

# ===== Build left menu from $Actions =====
$menuList.Items.Clear()
$Actions | ForEach-Object { [void]$menuList.Items.Add($_.Menu) }
$menuList.SelectedIndex = [int]$script:Settings.MenuIndex
$menuList.Add_SelectionChanged({ Refresh-Fields })
Refresh-Fields

# ===== Grid context menu =====
$cm = New-Object System.Windows.Controls.ContextMenu
$miCopy = New-Object System.Windows.Controls.MenuItem; $miCopy.Header = "Copy selected as CSV"
$miCopy.Add_Click({ if ($gridOut.SelectedItems -and $gridOut.SelectedItems.Count -gt 0) { $csv = $gridOut.SelectedItems | ConvertTo-Csv -NoTypeInformation | Out-String; [System.Windows.Clipboard]::SetText($csv.Trim()) } })
$cm.Items.Add($miCopy) | Out-Null
$gridOut.ContextMenu = $cm

# ===== Filter box =====
$txtFilter.Add_TextChanged({ $q = $txtFilter.Text; $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridOut.ItemsSource); if ($null -eq $view) { return }; if ([string]::IsNullOrWhiteSpace($q)) { $view.Filter = $null; return }; $view.Filter = { param($row) ($row.PSObject.Properties.Value | ForEach-Object { "$_" }) -join ' ' -match [regex]::Escape($q) } })

# ===== Export helpers =====
function Export-GridToCsv {
  if (-not $gridOut.ItemsSource) { return }
  $path = Join-Path "C:\Temp" ("ExAdminWpf_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
  try {
    $gridOut.ItemsSource | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    [System.Windows.Forms.MessageBox]::Show("Saved to `n$path","Export CSV") | Out-Null
  } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Export CSV") | Out-Null }
}
function Export-GridToXlsx {
  if (-not $gridOut.ItemsSource) { return }
  if (-not (Ensure-ExcelModule)) {
    [System.Windows.Forms.MessageBox]::Show("ImportExcel module not available","Export Excel") | Out-Null
    return
  }
  $path = Join-Path "C:\Temp" ("ExAdminWpf_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
  try {
    $gridOut.ItemsSource | Export-Excel -Path $path -AutoSize -WorksheetName 'Data' -Title 'Exchange Admin WPF Export'
    [System.Windows.Forms.MessageBox]::Show("Saved to `n$path","Export Excel") | Out-Null
  } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Export Excel") | Out-Null }
}
function Autosize-GridColumns { if ($gridOut.Columns) { foreach($c in $gridOut.Columns){ $c.Width = [System.Windows.Controls.DataGridLength]::Auto } } }

# ===== Menu & button wiring =====
$btnConnect.Add_Click({ Connect-EXO })
$btnDisconnect.Add_Click({ Disconnect-EXO })
$mnuConnect.Add_Click({ Connect-EXO })
$mnuDisconnect.Add_Click({ Disconnect-EXO })
$mnuExit.Add_Click({ $win.Close() })
$mnuRefresh.Add_Click({ Refresh-Fields })
$mnuClear.Add_Click({ $gridOut.ItemsSource = $null; $txtFilter.Text = ""; $lblStatus.Text = "Ready" })
$mnuAutosize.Add_Click({ Autosize-GridColumns })
$mnuAbout.Add_Click({ [System.Windows.Forms.MessageBox]::Show("Exchange Admin WPF`nBuild: $(Get-Date)","About") | Out-Null })
$mnuShortcuts.Add_Click({ [System.Windows.Forms.MessageBox]::Show("Alt+F/V/H = Menus`nEnter = Run, Esc = Close`nF5 = Refresh Fields`nCtrl+E = Export Excel, Ctrl+Shift+C = Export CSV`nCtrl+K = Connect, Ctrl+D = Disconnect","Shortcuts") | Out-Null })
$mnuExportCsv.Add_Click({ Export-GridToCsv })
$mnuExportXlsx.Add_Click({ Export-GridToXlsx })

$btnClear.Add_Click({ $gridOut.ItemsSource = $null; $txtFilter.Text = ""; $lblStatus.Text = "Ready" })
$btnCopy.Add_Click({ $sel = $gridOut.SelectedItems; if ($sel -and $sel.Count -gt 0) { $csv = $sel | ConvertTo-Csv -NoTypeInformation | Out-String; [System.Windows.Clipboard]::SetText($csv.Trim()); $lblStatus.Text = "Copied selection (CSV)" } else { $lblStatus.Text = "Nothing selected" } })

# ===== Dispatcher (Run) =====
$btnRun.Add_Click({
  try {
    $A = Get-CurrentAction; if (-not $A) { return }
    $ctx = [pscustomobject]@{
      Primary   = ($txtPrimary.Text).Trim()
      Secondary = ($txtSecondary.Text).Trim()
      Option    = if ($cboOption.SelectedItem){ $cboOption.SelectedItem.ToString() } else { "" }
      Extra     = ($txtExtra.Text).Trim()
      Start     = $txtStart.Text
      End       = $txtEnd.Text
      MsgInt    = $txtMsgInt.Text
      MsgExt    = $txtMsgExt.Text
    }
    # Only require Primary if the action's UI says it is needed
if ($A.UI.Primary -and [string]::IsNullOrWhiteSpace($ctx.Primary)) {
  Present-Data ([pscustomobject]@{ Info="Provide 'Mailbox / DL / PF Path'." })
  return
}

    if ($A.Heavy -and -not (Test-ExchangeOnline)) {
      Present-Data ([pscustomobject]@{ Info="Not connected to Exchange Online. Use File → Connect." })
      return
    }

    $useAsync = [bool]$script:Settings.UseAsyncForHeavy -and $A.Heavy  # default false
    if ($useAsync) {
      $sb = { param($run,$ctx) & $run $ctx }
      Invoke-Async $sb @{ run = $A.Run; ctx = $ctx }
      return
    }

    Set-Busy $true
    $result = & $A.Run $ctx
    Present-Data $result
  } catch {
    Present-Data ([pscustomobject]@{ Error = $_.Exception.Message })
  } finally {
    Set-Busy $false
  }
})

# Keyboard shortcuts
$win.Add_PreviewKeyDown({
  $ctrl = [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl) -or
          [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RightCtrl)
  $shift = [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftShift) -or
           [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RightShift)
  switch ($_.Key) {
    'Enter' { $btnRun.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
    'Escape'{ $win.Close() }
    'F5'    { Refresh-Fields }
    'E'     { if ($ctrl -and -not $shift) { Export-GridToXlsx } }
    'C'     { if ($ctrl -and  $shift)     { Export-GridToCsv } }
    'K'     { if ($ctrl) { Connect-EXO } }
    'D'     { if ($ctrl) { Disconnect-EXO } }
  }
})

# Status indicators
if (Test-ExchangeOnline) { $lblConn.Text="Connected"; $dotConn.Fill=[Windows.Media.Brushes]::LimeGreen } else { $lblConn.Text="Not connected"; $dotConn.Fill=[Windows.Media.Brushes]::Tomato }
if (Ensure-ExcelModule) { $lblExcel.Text="ImportExcel: available"; $dotExcel.Fill=[Windows.Media.Brushes]::LimeGreen } else { $lblExcel.Text="ImportExcel: missing"; $dotExcel.Fill=[Windows.Media.Brushes]::Tomato }

# Restore settings
if ($script:Settings.Primary)   { $txtPrimary.Text = $script:Settings.Primary }
if ($script:Settings.Secondary) { $txtSecondary.Text = $script:Settings.Secondary }

# Save settings on close
$win.Add_Closed({ Save-Settings })

# Show UI
[void]$win.ShowDialog()


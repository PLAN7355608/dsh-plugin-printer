# Shared WIA helpers, dot-sourced by scan.ps1, scanner-info.ps1 and list-scanners.ps1 so the device
# access code and the format tables exist in exactly one place.
#
# WIA scripting notes learned from real hardware:
#   * Properties.Item(n) takes a 1-based INDEX, not a property id, so lookups scan the list.
#   * The accepted resolution set is discovered by writing each candidate and reading it back:
#     drivers silently round the value, so the read-back is the only reliable answer.
#   * DPS_DOCUMENT_HANDLING_SELECT (3088) holds whatever the driver used last, so callers must not
#     trust it to tell whether the feeder is loaded.

$FORMAT_IDS = @{
  'bmp'  = '{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}'
  'jpg'  = '{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}'
  'jpeg' = '{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}'
  'png'  = '{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}'
  'tif'  = '{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}'
  'tiff' = '{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}'
}
$FORMAT_NAMES = @{}
foreach ($key in $FORMAT_IDS.Keys) { $FORMAT_NAMES[$FORMAT_IDS[$key]] = $key }

# The scanner tool's own resolution ceiling; the probe must not ask for more than the tool accepts.
$DPI_CANDIDATES = @(75, 100, 150, 200, 300, 400, 600, 800, 1200)

function Format-Name([string]$guid) {
  if (-not $guid) { return '' }
  $up = $guid.ToUpperInvariant()
  if ($FORMAT_NAMES.ContainsKey($up)) {
    $n = [string]$FORMAT_NAMES[$up]
    if ($n -eq 'jpeg') { return 'jpg' }
    if ($n -eq 'tiff') { return 'tif' }
    return $n
  }
  return ''
}

function Find-Property($properties, [int]$id) {
  $count = 0
  try { $count = [int]$properties.Count } catch { $count = 0 }
  for ($i = 1; $i -le $count; $i++) {
    try {
      $p = $properties.Item($i)
      if ([int]$p.PropertyID -eq $id) { return $p }
    } catch { }
  }
  return $null
}

function Get-PropertyById($properties, [int]$id) {
  $p = Find-Property $properties $id
  if ($p) {
    try { return $p.Value } catch { }
  }
  return $null
}

function Set-PropertyById($properties, [int]$id, $value) {
  $p = Find-Property $properties $id
  if ($p) {
    try { $p.Value = $value; return $true } catch { }
  }
  return $false
}

function Get-PropertyByName($properties, [string]$name) {
  $count = 0
  try { $count = [int]$properties.Count } catch { $count = 0 }
  for ($i = 1; $i -le $count; $i++) {
    try {
      $p = $properties.Item($i)
      if ([string]$p.Name -eq $name) { return [string]$p.Value }
    } catch { }
  }
  return ''
}

# Formats the device says it can deliver; unusable or unknown entries are dropped. Many drivers
# report nothing at all, so an empty list is normal, not an error.
function Get-SupportedFormats($properties) {
  $list = @()
  try {
    $p = Find-Property $properties 4106
    if ($p) {
      $st = $p.SubType
      $n = 0
      try { $n = [int]$st.Count } catch { $n = 0 }
      for ($i = 1; $i -le $n; $i++) {
        $g = ''
        try { $g = [string]$st.Item($i) } catch { try { $g = [string]$st[$i - 1] } catch { $g = '' } }
        if ($g -notmatch '^[{][0-9A-Fa-f]{8}-') { continue }
        $name = Format-Name $g
        if ($name -and ($list -notcontains $name)) { $list += $name }
      }
    }
  } catch { }
  return $list
}

# The accepted set is discovered by writing each candidate and reading it back.
function Get-AcceptedResolutions($item) {
  $accepted = @()
  foreach ($v in $DPI_CANDIDATES) {
    try {
      if (Set-PropertyById $item.Properties 6147 $v) {
        Set-PropertyById $item.Properties 6148 $v | Out-Null
        $back = Get-PropertyById $item.Properties 6147
        if ($null -ne $back -and [int]$back -eq $v) { $accepted += $v }
      }
    } catch { }
  }
  return $accepted
}

# Opens WIA and returns the device list. ok=false means WIA itself is unusable (no WIA service, COM
# refused, ...), which must be reported as an error and not as "no scanner is installed".
function Get-WiaDevices {
  $result = @{ ok = $false; error = ''; devices = @(); manager = $null; infos = $null }
  try {
    $dm = New-Object -ComObject WIA.DeviceManager
    $infos = $dm.DeviceInfos
    $count = 0
    try { $count = [int]$infos.Count } catch { $count = 0 }
    $list = @()
    for ($i = 1; $i -le $count; $i++) {
      $info = $infos.Item($i)
      $id = ''
      try { $id = [string]$info.DeviceID } catch { }
      $type = ''
      try { $type = [string]$info.Type } catch { }
      $name = Get-PropertyByName $info.Properties 'Name'
      if (-not $name) { $name = 'WIA device ' + $i }
      $list += @{ index = $i; name = $name; id = $id; type = $type; info = $info }
    }
    $result.ok = $true
    $result.devices = $list
    $result.manager = $dm
    $result.infos = $infos
  } catch {
    $result.error = $_.Exception.Message
  }
  return $result
}

# "scanner not found: X" without the available names forces a second tool call.
function Get-WiaNotFoundText([string]$wanted, $devices) {
  $names = @()
  foreach ($d in $devices) { $names += [string]$d.name }
  if ($names.Count -eq 0) { return ('scanner not found: ' + $wanted + ' (no WIA scanner is listed on this computer)') }
  return ('scanner not found: ' + $wanted + '; available scanners: ' + ($names -join ', '))
}

# Tries to open one device. A stale entry (the eSCL pseudo-device some all-in-one drivers leave
# next to the real WIA device) fails here, which lets the callers tell "installed but not usable"
# apart from "not installed" instead of offering a device that cannot be opened.
function Open-WiaDevice($device) {
  $result = @{ ok = $false; device = $null; error = '' }
  try {
    $result.device = $device.info.Connect()
    $result.ok = $true
  } catch {
    $result.error = $_.Exception.Message
  }
  return $result
}

# COM objects are not garbage-collected deterministically: a scanner stays locked until they go.
function Release-ComObject($obj) {
  if ($null -eq $obj) { return }
  try {
    if ([System.Runtime.InteropServices.Marshal]::IsComObject($obj)) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
  } catch { }
}

# Device lookup by exact name/id/index first, then by unique name prefix. The printer tools accept a
# prefix, so the scan tools should too: retyping "(HP Color LaserJet MFP M283cdw)" every time is the
# kind of friction that makes a working tool look broken.
function Find-WiaDevice($devices, [string]$wanted) {
  $result = @{ device = $null; ambiguous = @(); names = @() }
  if (-not $wanted) { return $result }
  foreach ($d in $devices) {
    if ($wanted -eq [string]$d.name -or $wanted -eq [string]$d.id -or $wanted -eq [string]$d.index) {
      $result.device = $d
      return $result
    }
  }
  $hits = @()
  foreach ($d in $devices) {
    $name = [string]$d.name
    if ($name.StartsWith($wanted, [System.StringComparison]::OrdinalIgnoreCase)) { $hits += $d }
  }
  if ($hits.Count -eq 1) { $result.device = $hits[0]; return $result }
  if ($hits.Count -gt 1) {
    $result.ambiguous = $hits
    foreach ($h in $hits) { $result.names += [string]$h.name }
  }
  return $result
}

# A4 in pixels at the negotiated resolution (210 x 297 mm). This HP WSD entry reports a Legal
# (8.5 x 14 in) scan area for its feeder, so every page comes back with ~76 mm of blank tail.
function Get-A4Extent([int]$dpi) {
  if ($dpi -le 0) { return $null }
  return @{
    width = [int][System.Math]::Round(210.0 / 25.4 * $dpi)
    height = [int][System.Math]::Round(297.0 / 25.4 * $dpi)
  }
}

# Asks the device for an A4 scan area before the transfer. Best effort: drivers that do not expose
# the extent properties keep their own area, and crop=a4 trims the result afterwards anyway.
function Set-WiaA4Extent($item, [int]$dpi) {
  $extent = Get-A4Extent $dpi
  if (-not $extent) { return $false }
  $ok = $true
  try {
    Set-PropertyById $item.Properties 6149 0 | Out-Null
    Set-PropertyById $item.Properties 6150 0 | Out-Null
    if (-not (Set-PropertyById $item.Properties 6151 $extent.width)) { $ok = $false }
    if (-not (Set-PropertyById $item.Properties 6152 $extent.height)) { $ok = $false }
  } catch { $ok = $false }
  return $ok
}

# DPS_PAGES (3096) is deliberately left alone. Setting it to 0 ("keep feeding until the tray is
# empty") is the textbook way to do multi-page ADF scans, but this HP WSD driver answers every
# transfer with E_INVALIDARG afterwards, so multi-page scanning opens one WIA session per sheet
# instead (see Open-ScanItem in scan.ps1) and relies on the PAPER_EMPTY status to end the run.

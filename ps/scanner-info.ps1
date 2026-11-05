param(
  [string]$Device = ''
)
# Reports what each WIA scanner can do without scanning anything: document sources, the formats and
# resolutions the driver really accepts, and the current document-handling settings.
#   -Device  only describe this device (name, DeviceID, or 1-based index); empty = all devices
# Records (US = 0x1f):
#   NONE    <US> (no WIA scanner is installed)
#   SCANNER <US> name <US> type <US> id
#   DETAIL  <US> name <US> caps <US> formats <US> resolutions <US> source <US> docStatus <US> note
#   CODE    <US> scannerUnreadable | scannerNotFound   (always immediately followed by ERR)
#   ERR     <US> message
# A WIA failure is an error; "no scanner is installed" is reported as NONE and is not one.
. (Join-Path $PSScriptRoot 'wia-common.ps1')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$US = [char]0x1f
function Emit([string[]]$fields) { [Console]::Out.WriteLine([string]::Join($US, $fields)) }
function Fail([string]$message, [string]$code) {
  if ($code) { Emit @('CODE', $code) }
  Emit @('ERR', $message)
  exit 1
}

$wia = $null
$dev = $null
$item = $null
try {
  $wia = Get-WiaDevices
  if (-not $wia.ok) { Fail ('cannot open the WIA device manager: ' + $wia.error) 'scannerUnreadable' }
  if (@($wia.devices).Count -eq 0) {
    Emit @('NONE', 'no WIA scanner is installed on this computer')
    exit 0
  }
  $want = ''
  if ($Device) { $want = [string]$Device }
  $matched = 0
  foreach ($d in $wia.devices) {
    if ($want -and $want -ne [string]$d.name -and $want -ne [string]$d.id -and $want -ne [string]$d.index) { continue }
    $matched++
    $name = [string]$d.name
    Emit @('SCANNER', $name, [string]$d.type, [string]$d.id)

    $caps = ''
    $formats = ''
    $resolutions = ''
    $source = ''
    $docStatus = ''
    $note = ''
    $dev = $null
    try { $dev = $d.info.Connect() } catch { $note = 'could not open the device: ' + $_.Exception.Message }
    if ($dev) {
      try {
        $raw = Get-PropertyById $dev.Properties 3086
        if ($null -ne $raw) { $caps = [string]([int]$raw) }
      } catch { }
      try {
        $raw = Get-PropertyById $dev.Properties 3088
        if ($null -ne $raw) { $source = [string]([int]$raw) }
      } catch { }
      try {
        $raw = Get-PropertyById $dev.Properties 3087
        if ($null -ne $raw) { $docStatus = [string]([int]$raw) }
      } catch { }
      try {
        $items = 0
        try { $items = [int]$dev.Items.Count } catch { $items = 0 }
        if ($items -gt 0) {
          $item = $dev.Items.Item(1)
          $formats = [string]::Join(',', @(Get-SupportedFormats $item.Properties))
          $resolutions = [string]::Join(',', @(Get-AcceptedResolutions $item))
          if ($formats -eq '') { $note = 'the driver does not publish its format list' }
        } else {
          $note = 'no scan item is available (nothing on the glass, or the feeder is empty)'
        }
      } catch { $note = 'could not read the device items: ' + $_.Exception.Message }
    }
    Emit @('DETAIL', $name, $caps, $formats, $resolutions, $source, $docStatus, $note)
  }
  if ($want -and $matched -eq 0) { Fail (Get-WiaNotFoundText $want $wia.devices) 'scannerNotFound' }
} catch {
  Emit @('ERR', $_.Exception.Message)
  exit 1
} finally {
  if ($item) { Release-ComObject $item }
  if ($dev) { Release-ComObject $dev }
  if ($wia) {
    if ($wia.infos) { Release-ComObject $wia.infos }
    if ($wia.manager) { Release-ComObject $wia.manager }
  }
}

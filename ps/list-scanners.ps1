param()
# Lists WIA scanners (a scanner driver also exposes the flatbed/feeder as a WIA device).
# Output (US = 0x1f):
#   SCANNER <US> Name <US> Type <US> DeviceID <US> available <US> error | WARN ... | DONE
#                     CODE <US> scannerUnreadable | ERR <US> message
# WIA being unavailable is an error, not "no scanner is installed": the two look identical in an
# empty list, so they are reported differently.
#
# Every device is opened briefly to find out whether it is usable right now. All-in-one drivers
# often leave a second, permanent off-line eSCL entry next to the real scanner; without this probe
# both look identical in the list and a caller can pick the one that never works.
. (Join-Path $PSScriptRoot 'wia-common.ps1')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$US = [char]0x1f
function Emit([string[]]$fields) { [Console]::Out.WriteLine([string]::Join($US, $fields)) }

$wia = $null
$dev = $null
try {
  $wia = Get-WiaDevices
  if (-not $wia.ok) {
    Emit @('CODE', 'scannerUnreadable')
    Emit @('ERR', ('cannot open the WIA device manager: ' + $wia.error))
    exit 1
  }
  foreach ($d in $wia.devices) {
    $probe = Open-WiaDevice $d
    if ($probe.device) { $dev = $probe.device; Release-ComObject $dev; $dev = $null }
    $available = '0'
    if ($probe.ok) { $available = '1' }
    Emit @('SCANNER', [string]$d.name, [string]$d.type, [string]$d.id, $available, [string]$probe.error)
  }
  Emit @('DONE')
} catch {
  Emit @('CODE', 'scannerUnreadable')
  Emit @('ERR', $_.Exception.Message)
  exit 1
} finally {
  if ($dev) { Release-ComObject $dev }
  if ($wia) {
    if ($wia.infos) { Release-ComObject $wia.infos }
    if ($wia.manager) { Release-ComObject $wia.manager }
  }
}

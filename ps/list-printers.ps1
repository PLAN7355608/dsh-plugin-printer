param()
# Lists installed printers and the current default printer.
#
# Two sources: the registry (always readable, even by a low-integrity process) and WMI. The registry
# is the primary source for names, drivers and ports; WMI enriches live status and queue length and
# is the fallback when the registry could not be read. A failure of one source is reported as a WARN;
# when BOTH sources fail the tool reports CODE listUnreadable and exits non-zero, because an empty
# list would otherwise look exactly like "this machine has no printers".
#
# Live status is deliberately reported ONLY from WMI: the printed-queue "Status" value in the registry
# is not the Win32 PRINTER_STATUS_* bitmask in practice (on this machine all six queues, virtual ones
# included, carry the same 0x0180), so decoding those bits labels every printer "Offline". The bit
# tables live in print-common.ps1, shared with printer-action.ps1.
#
# Output: one record per line, fields separated by US (\x1f).
#   PRINTER <US> Name <US> Driver <US> Port <US> Status <US> Jobs <US> Shared <US> Note <US> Source
#   DEFAULT <US> DefaultPrinterName
#   WARN <US> list-printers <US> registry|wmi <US> message
#   CODE <US> listUnreadable
#   ERR <US> message
#   DONE
# Output is UTF-8; all display text is produced by the host, never by this script.
. (Join-Path $PSScriptRoot 'print-common.ps1')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$US = [char]0x1f
function Emit([string[]]$fields) { [Console]::Out.WriteLine([string]::Join($US, $fields)) }

try {
  $defaultName = Get-DefaultPrinterName

  $wmiError = ''
  $wmiByName = @{}
  try {
    foreach ($p in @(Get-WmiObject -Class Win32_Printer -ErrorAction Stop)) {
      $n = [string]$p.Name
      if ($n) { $wmiByName[$n] = $p }
    }
  } catch {
    $wmiError = Get-ErrorText $_
    Emit @('WARN', 'list-printers', 'wmi', $wmiError)
  }

  $root = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
  $regError = ''
  $printerNames = @()
  try {
    $printerNames = @(Get-ChildItem -Path $root -ErrorAction Stop | Select-Object -ExpandProperty PSChildName)
  } catch {
    $regError = Get-ErrorText $_
    Emit @('WARN', 'list-printers', 'registry', $regError)
  }

  if ($printerNames.Count -eq 0) {
    foreach ($key in $wmiByName.Keys) { $printerNames += $key }
  }

  if ($regError -and $wmiError) {
    Emit @('CODE', 'listUnreadable')
    Emit @('ERR', ('neither the registry printer list nor WMI could be read: registry=' + $regError + '; wmi=' + $wmiError))
    exit 1
  }

  foreach ($name in ($printerNames | Sort-Object)) {
    $driver = ''
    $port = ''
    $shared = 'False'
    $note = ''
    $status = ''
    $jobs = ''
    $source = 'Registry'
    try {
      $k = Get-ItemProperty -Path (Join-Path $root $name) -ErrorAction Stop
      if ($k.'Printer Driver') { $driver = [string]$k.'Printer Driver' }
      if (-not $driver -and $k.Driver) { $driver = [string]$k.Driver }
      if ($k.Port) { $port = [string]$k.Port }
      if ($k.'Share Name') { $shared = 'True' }
      try { if ((([int]$k.Attributes) -band 8) -eq 8) { $shared = 'True' } } catch { }
    } catch { }

    if ($wmiByName.ContainsKey($name)) {
      $source = 'WMI'
      $p = $wmiByName[$name]
      if (-not $driver) { $driver = [string]$p.DriverName }
      if (-not $port) { $port = [string]$p.PortName }
      if ($null -ne $p.Jobs) { $jobs = [string]$p.Jobs }
      if ($p.Shared) { $shared = 'True' }
      if ($p.Default) { $note = 'default' }
      $status = Get-PrinterStatusText $p
    }

    if ($defaultName -and $name -eq $defaultName) { $note = 'default' }
    Emit @('PRINTER', $name, $driver, $port, $status, $jobs, $shared, $note, $source)
  }

  Emit @('DEFAULT', $defaultName)
  Emit @('DONE')
} catch {
  Emit @('CODE', 'listUnreadable')
  Emit @('ERR', (Get-ErrorText $_))
  exit 1
}

# Shared helpers for the printer scripts, dot-sourced by print-file.ps1, list-printers.ps1 and
# printer-action.ps1 so that the registry probing and the status decoding exist in exactly one place.
#
# Nothing here has side effects: no spooler write, no printer preference change.

# --- shell association probing -------------------------------------------------------------
# The registry, not the shell "assoc" command: assoc resolves through the same UserChoice layer but
# reports nothing for extensions whose association lives in HKCU only.
function Get-RegDefault([string]$key) {
  try {
    $item = Get-Item -LiteralPath $key -ErrorAction Stop
    $value = $item.GetValue('')
    if ($null -eq $value) { return '' }
    return [string]$value
  } catch { return '' }
}

function Get-AssocProgId([string]$path) {
  $ext = [System.IO.Path]::GetExtension($path)
  if (-not $ext) { return '' }
  return (Get-RegDefault ('Registry::HKEY_CLASSES_ROOT\' + $ext))
}

# Returns @{ verb = 'printto' | 'print' | ''; command = '...'; progId = '...' } for this file type.
function Get-PrintVerb([string]$path) {
  $ext = [System.IO.Path]::GetExtension($path)
  $progId = Get-AssocProgId $path
  $targets = @()
  # NOTE: a ProviderPath ("Registry::HKEY_CLASSES_ROOT...") needs a backslash before every child key;
  # without them the whole path is one unknown key name, Get-RegDefault swallows the error and every
  # file type looks like it has no print verb.
  if ($progId) { $targets += ('Registry::HKEY_CLASSES_ROOT\' + $progId) }
  if ($ext) { $targets += ('Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\' + $ext) }
  foreach ($verb in @('printto', 'print')) {
    foreach ($t in $targets) {
      $cmd = Get-RegDefault ($t + '\shell\' + $verb + '\command')
      if ($cmd) { return @{ verb = $verb; command = $cmd; progId = $progId } }
    }
  }
  return @{ verb = ''; command = ''; progId = $progId }
}

# WPS Office registers a print verb but drops the printer argument on the floor, so a "successful"
# print lands on whatever the system default printer is. Detecting it lets the tool refuse loudly.
function Test-PrinterArgIgnored([string]$command) {
  if (-not $command) { return $false }
  if ($command -match '(?i)kingsoft|[\\/]wps|kwps|ket\.exe|et\.exe|wpp\.exe|wpspdf') { return $true }
  return $false
}

# --- installed printers --------------------------------------------------------------------
# Registry first; WMI is the fallback when the registry key cannot be read.
function Get-InstalledPrinterNames {
  $names = @()
  try {
    $names = @(Get-ChildItem -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -ErrorAction Stop | Select-Object -ExpandProperty PSChildName)
  } catch { $names = @() }
  if ($names.Count -eq 0) {
    try {
      $names = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | ForEach-Object { [string]$_.Name })
    } catch { $names = @() }
  }
  return $names
}

# A "printer not found" message without the available names forces the user to run another tool.
function Get-PrinterNotFoundText([string]$printer) {
  $names = @(Get-InstalledPrinterNames)
  if ($names.Count -eq 0) { return ('printer not found: ' + $printer + ' (the installed printer list could not be read either)') }
  return ('printer not found: ' + $printer + '; available printers: ' + (($names | Sort-Object) -join ', '))
}

# --- system default printer ----------------------------------------------------------------
# The registry value is "Printer Name,winspool,port". The name itself may contain a comma, so it is
# split at the LAST comma, not the first (the two splitting styles had drifted apart between scripts).
function Get-DefaultPrinterName {
  $name = Get-RegDefault 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows\Device'
  if ($name) {
    $idx = $name.LastIndexOf(',')
    if ($idx -gt 0) { return $name.Substring(0, $idx) }
    return $name
  }
  try {
    $p = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | Where-Object { $_.Default -eq $true } | Select-Object -First 1
    if ($p) { return [string]$p.Name }
  } catch { }
  return ''
}

# --- error text ----------------------------------------------------------------------------
# A caught error must never come out as an empty string: an empty message silently turned a failed
# duplex change into "duplex off" with no warning at all (caught on a real HP M283cdw).
function Get-ErrorText($err) {
  $parts = @()
  if ($err.Exception -and $err.Exception.Message) { $parts += [string]$err.Exception.Message }
  if ($err.Exception.InnerException -and $err.Exception.InnerException.Message) { $parts += [string]$err.Exception.InnerException.Message }
  if (-not $parts.Count -and $err.CategoryInfo -and $err.CategoryInfo.Reason) { $parts += [string]$err.CategoryInfo.Reason }
  if (-not $parts.Count -and $err.FullyQualifiedErrorId) { $parts += [string]$err.FullyQualifiedErrorId }
  if (-not $parts.Count) { $parts += 'unknown failure' }
  return ($parts -join ' / ')
}

# --- status decoding -----------------------------------------------------------------------
# Win32_Printer.PrinterStatus: 1=Other 2=Unknown 3=Idle 4=Printing 5=Warmup 6=Stopped 7=Offline
$script:PrinterStatusWords = @{ '3' = 'Idle'; '4' = 'Printing'; '5' = 'Warmup'; '6' = 'Stopped'; '7' = 'Offline' }

# PRINTER_STATUS_* bits that matter for troubleshooting. PrinterState is the live bitmask.
$script:PrinterStateBits = @(
  @(0x2, 'Error'), @(0x8, 'PaperJam'), @(0x10, 'PaperOut'), @(0x20, 'ManualFeed'),
  @(0x40, 'PaperProblem'), @(0x80, 'Offline'), @(0x800, 'OutputBinFull'), @(0x1000, 'NotAvailable'),
  @(0x20000, 'LowToner'), @(0x40000, 'NoToner'), @(0x100000, 'UserIntervention'),
  @(0x200000, 'OutOfMemory'), @(0x400000, 'DoorOpen'), @(0x800000, 'ServerUnknown')
)

# DetectedErrorState adds conditions that have no PRINTER_STATUS_* bit of their own.
$script:DetectedErrorStates = @{
  '3' = 'LowPaper'; '4' = 'NoPaper'; '5' = 'LowToner'; '6' = 'NoToner'; '7' = 'DoorOpen'
  '8' = 'PaperJam'; '9' = 'Offline'; '10' = 'ServiceRequested'; '11' = 'OutputBinFull'
  '12' = 'PaperProblem'; '13' = 'CannotPrintPage'; '14' = 'UserIntervention'; '15' = 'OutOfMemory'
  '16' = 'ServerUnknown'
}

function Get-PrinterStatusWord([string]$printerStatus) {
  if ($script:PrinterStatusWords.ContainsKey($printerStatus)) { return [string]$script:PrinterStatusWords[$printerStatus] }
  return ''
}

# A printer out of toner is still "idle": the conditions are returned separately so a caller can
# report them even when the plain status word looks healthy.
function Get-PrinterConditionList($printer) {
  $conds = @()
  if (-not $printer) { return $conds }
  $bits = 0
  try { $bits = [int]$printer.PrinterState } catch { $bits = 0 }
  foreach ($pair in $script:PrinterStateBits) {
    if (($bits -band [int]$pair[0]) -ne 0) { $conds += [string]$pair[1] }
  }
  $des = ''
  try { $des = [string]$printer.DetectedErrorState } catch { $des = '' }
  if ($des -and $script:DetectedErrorStates.ContainsKey($des) -and ($conds -notcontains $script:DetectedErrorStates[$des])) {
    $conds += [string]$script:DetectedErrorStates[$des]
  }
  $offline = $false
  try { $offline = [bool]$printer.WorkOffline } catch { $offline = $false }
  if ($offline -and ($conds -notcontains 'Offline')) { $conds += 'Offline' }
  return $conds
}

# "Idle+PaperOut", "Printing", ... - one string for both the list and the test-page result.
function Get-PrinterStatusText($printer) {
  $word = ''
  if ($printer) {
    try { $word = Get-PrinterStatusWord ([string]$printer.PrinterStatus) } catch { $word = '' }
  }
  $conds = @(Get-PrinterConditionList $printer)
  if ($conds.Count -eq 0) { return $word }
  $joined = $conds -join '+'
  if ($word) { return $word + '+' + $joined }
  return $joined
}

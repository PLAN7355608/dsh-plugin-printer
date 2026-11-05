param(
  [Parameter(Mandatory = $true)][string]$Printer
)
# Hardware capabilities of one printer, for printer_info.
# Records (US = 0x1f):
#   CAPS   <US> name <US> canDuplex <US> duplexOn <US> color <US> xres <US> yres <US> papers <US> comment <US> location <US> duplexSource <US> colorSource
#   CONFIG <US> duplexing <US> color <US> paperSize
#   WARN   <US> kind <US> text
#   CODE   <US> code   (always immediately followed by ERR)
#   ERR    <US> message
#
# Field sources, in priority order:
#   1. the WMI Win32_Printer properties Duplex / Color (WMI may not expose them at all);
#   2. the Win32_Printer Capabilities bitmask (2 = colour printing, 4 = duplex printing);
#   3. the GDI printer driver (System.Drawing PrinterSettings.CanDuplex / SupportsColor).
# 1.3.0 only read source 1. This machine's HP V4 driver leaves both properties empty while
# Capabilities still carries 2 and 4, so printer_info used to report "double-sided: unknown,
# colour: unknown" for a printer that duplexes and prints colour. The source of every value is
# carried in the CAPS record so the host can say "inferred" instead of presenting it as a driver fact.
#
# A WMI query that is denied must NOT look like "a printer with no capabilities": that case exits
# non-zero with CODE capsUnreadable; an unknown name gets CODE unknownPrinter.
#
# duplexOn is derived from the printing preferences (DuplexingMode), not from the capability flag:
# the flag only says duplexing is *allowed*, the preference says whether it is currently on.
#
# Get-PrintConfiguration needs Windows 8+; on Windows 7 the CONFIG record is missing and a WARN is
# emitted instead of silently reporting empty preferences.
. (Join-Path $PSScriptRoot 'print-common.ps1')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$US = [char]0x1f
function Emit([string[]]$fields) { [Console]::Out.WriteLine([string]::Join($US, $fields)) }
function Fail([string]$message, [string]$code) {
  if ($code) { Emit @('CODE', $code) }
  Emit @('ERR', $message)
  exit 1
}

function BoolText($value) {
  if ($null -eq $value) { return '' }
  if ([bool]$value) { return '1' }
  return '0'
}

$installed = @(Get-InstalledPrinterNames)
$all = $null
$queryError = ''
try {
  $all = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)
} catch {
  try {
    $all = @(Get-WmiObject -Class Win32_Printer -ErrorAction Stop)
  } catch {
    $all = $null
    $queryError = Get-ErrorText $_
  }
}

if (-not $all) {
  if ($installed.Count -gt 0 -and ($installed -notcontains $Printer)) {
    Fail (Get-PrinterNotFoundText $Printer) 'unknownPrinter'
  }
  Fail ('the printer capabilities could not be read: Win32_Printer is not accessible (' + $queryError + '). This does not mean the printer has no capabilities - the query was refused, usually because the process lacks the required privilege.') 'capsUnreadable'
}

$p = $all | Where-Object { $_.Name -eq $Printer } | Select-Object -First 1
if (-not $p) { Fail (Get-PrinterNotFoundText $Printer) 'unknownPrinter' }

$papers = ''
try {
  if ($p.PaperSizesSupported) { $papers = [string]::Join(',', @($p.PaperSizesSupported)) }
} catch { }

# Printing preferences first: duplexOn is derived from the duplexing mode the driver will actually
# use, so the CAPS record can carry it.
$duplexing = ''
$colorPref = ''
$paper = ''
$configError = ''
try {
  if (-not (Get-Command -Name 'Get-PrintConfiguration' -ErrorAction SilentlyContinue)) {
    $configError = 'the printing preferences cannot be read on this Windows version (Get-PrintConfiguration needs Windows 8 or later)'
  } else {
    $cfg = Get-PrintConfiguration -PrinterName $Printer -ErrorAction Stop
    $duplexing = [string]$cfg.DuplexingMode
    try { $colorPref = [string]$cfg.Color } catch { }
    try { $paper = [string]$cfg.PaperSize } catch { }
  }
} catch {
  $configError = 'could not read the printing preferences: ' + (Get-ErrorText $_)
}

$duplexOn = ''
if ($duplexing) {
  if ($duplexing -match '(?i)simplex|onesided') { $duplexOn = '0' } else { $duplexOn = '1' }
}

# Win32_Printer.Capabilities: 2 = colour printing, 4 = duplex printing. Some drivers return only
# these bits and leave the plain boolean properties empty.
$capBits = 0
try {
  foreach ($bit in @($p.Capabilities)) { $capBits = $capBits -bor [int]$bit }
} catch { $capBits = 0 }

# GDI is the last resort: the same driver that renders the pages can answer whether it duplexes.
$gdi = $null
try {
  Add-Type -AssemblyName System.Drawing -ErrorAction Stop
  $settings = New-Object System.Drawing.Printing.PrinterSettings
  $settings.PrinterName = $Printer
  if ($settings.IsValid) { $gdi = $settings }
} catch { $gdi = $null }

$canDuplex = $null
$duplexSource = ''
try { if ($null -ne $p.Duplex) { $canDuplex = [bool]$p.Duplex; $duplexSource = 'wmi' } } catch { }
if ($null -eq $canDuplex -and (($capBits -band 4) -ne 0)) { $canDuplex = $true; $duplexSource = 'capabilities' }
if ($null -eq $canDuplex -and $gdi) { $canDuplex = [bool]$gdi.CanDuplex; $duplexSource = 'gdi' }

$color = $null
$colorSource = ''
try { if ($null -ne $p.Color) { $color = [bool]$p.Color; $colorSource = 'wmi' } } catch { }
if ($null -eq $color -and (($capBits -band 2) -ne 0)) { $color = $true; $colorSource = 'capabilities' }
if ($null -eq $color -and $gdi) { $color = [bool]$gdi.SupportsColor; $colorSource = 'gdi' }

Emit @('CAPS', [string]$p.Name, (BoolText $canDuplex), $duplexOn, (BoolText $color), [string]$p.HorizontalResolution, [string]$p.VerticalResolution, $papers, [string]$p.Comment, [string]$p.Location, $duplexSource, $colorSource)

if ($configError) {
  Emit @('WARN', 'config', $configError)
} else {
  Emit @('CONFIG', $duplexing, $colorPref, $paper)
}

param(
  [Parameter(Mandatory = $true)][string]$Action,
  [string]$Printer = ''
)
# Printer actions that do not need a document.
#   -Action set-default  -Printer <name>   set the system default printer
#   -Action test-page    -Printer <name>   print a Windows test page
#   -Action show-queue   -Printer <name>   open the printer queue window on screen
#
# set-default tries three mechanisms in turn (Win32 SetDefaultPrinter, the WScript.Network COM
# object, rundll32 printui /y) and then READS THE RESULT BACK from the registry. A silent no-op must
# not be reported as success: when the choice does not stick it emits WARN with the printer that is
# actually default. Windows 10/11 users should also know that with "Let Windows manage my default
# printer" enabled the default follows the last printer used, so any choice can be overridden by the
# next print job - the host decides how to phrase that.
#
# Every action validates the printer name first and, when it is unknown, reports the available
# printers instead of only saying "not found".
#
# Output records: OK <US> <action> <US> <printer> [<US> <detail> ...]
#                 WARN <US> <action> <US> <requested> <US> <actual>
#                 CODE <US> <code>   (always immediately followed by ERR)
#                 ERR <US> <message>
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

# First job in a printer's queue, or $null. Get-PrintJob needs Windows 8+, so Windows 7 falls back to
# the Win32_PrintJob class (job Name is "Printer,Id").
function Get-FirstJob([string]$name) {
  try {
    $jobs = @(Get-PrintJob -PrinterName $name -ErrorAction Stop)
    if ($jobs.Count -gt 0) { return $jobs[0] }
  } catch { }
  try {
    $jobs = @(Get-WmiObject -Class Win32_PrintJob -ErrorAction Stop | Where-Object { ([string]$_.Name) -like ($name + ',*') })
    if ($jobs.Count -gt 0) { return $jobs[0] }
  } catch { }
  return $null
}

function Invoke-SetDefaultPrinterApi([string]$name) {
  try {
    if (-not ('DshPrintNative' -as [type])) {
      Add-Type -Namespace DshPrint -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("winspool.drv", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
public static extern bool SetDefaultPrinter(string pszPrinter);
'@ -ErrorAction Stop
    }
    return [DshPrint.Native]::SetDefaultPrinter($name)
  } catch {
    return $false
  }
}

# printui is handed ONE command line: Start-Process joins an argument array with spaces without
# quoting the items, so a printer name containing spaces used to be split apart.
function Invoke-PrintUi([string]$verb, [string]$name) {
  $line = 'printui.dll,PrintUIEntry ' + $verb + ' /n "' + ($name -replace '"', '') + '"'
  Start-Process -FilePath 'rundll32.exe' -ArgumentList $line -Wait -WindowStyle Hidden | Out-Null
}

if (-not $Printer) { Fail 'no printer name given' 'noPrinter' }

try {
  $installed = @(Get-InstalledPrinterNames)
  if ($installed.Count -gt 0 -and ($installed -notcontains $Printer)) {
    Fail (Get-PrinterNotFoundText $Printer) 'unknownPrinter'
  }

  switch ($Action) {
    'set-default' {
      $method = ''
      $current = Get-DefaultPrinterName

      if ($current -ne $Printer) {
        Invoke-SetDefaultPrinterApi $Printer | Out-Null
        $current = Get-DefaultPrinterName
        if ($current -eq $Printer) { $method = 'api' }
      } else {
        $method = 'already'
      }

      if ($current -ne $Printer) {
        try {
          $net = New-Object -ComObject WScript.Network
          $net.SetDefaultPrinter($Printer)
          $current = Get-DefaultPrinterName
          if ($current -eq $Printer) { $method = 'com' }
        } catch { }
      }

      if ($current -ne $Printer) {
        try {
          Invoke-PrintUi '/y' $Printer
          $current = Get-DefaultPrinterName
          if ($current -eq $Printer) { $method = 'printui' }
        } catch { }
      }

      if ($current -ne $Printer) {
        Emit @('WARN', 'set-default', $Printer, $current)
      } else {
        Emit @('OK', 'set-default', $Printer, $current, $method)
      }
    }
    'test-page' {
      # printui /k only launches the print-UI verb: it returns 0 without necessarily creating any
      # spool job, so its exit code proves nothing. Use the WMI method (real Win32 return code),
      # report the device state we can read, and then watch the queue so the host can tell the
      # difference between "the spooler took the job" and "the printer produced paper".
      $all = $null
      try { $all = @(Get-WmiObject -Class Win32_Printer -ErrorAction Stop) } catch { $all = $null }
      if (-not $all) {
        # Restricted token: WMI is unreadable. Fall back to the print-UI verb and mark the result as
        # unverified instead of pretending it worked.
        Invoke-PrintUi '/k' $Printer
        Emit @('WARN', 'test-page', $Printer, '', '', 'unverified')
        break
      }
      $wmi = $all | Where-Object { $_.Name -eq $Printer } | Select-Object -First 1
      if (-not $wmi) { Fail (Get-PrinterNotFoundText $Printer) 'unknownPrinter' }

      $state = Get-PrinterStatusText $wmi
      $rv = ''
      try { $rv = [string]($wmi.PrintTestPage()).ReturnValue } catch { $rv = 'error: ' + (Get-ErrorText $_) }

      $jobStatus = ''
      $deadline = (Get-Date).AddSeconds(15)
      while ((Get-Date) -lt $deadline) {
        $job = Get-FirstJob $Printer
        if ($job) {
          $s = ''
          try { $s = [string]$job.JobStatus } catch { }
          if ($s) { $jobStatus = $s }
          if ($s -match 'Error|Offline|Blocked|Paper|Jam|User|Delet') { break }
        }
        Start-Sleep -Milliseconds 700
      }
      # No job was ever observed: report that honestly instead of an empty status that reads like
      # "printed, nothing to say".
      if (-not $jobStatus) {
        Emit @('WARN', 'test-page', $Printer, $rv, $state, 'unverified')
        break
      }
      Emit @('OK', 'test-page', $Printer, $rv, $state, $jobStatus)
    }
    'show-queue' {
      Invoke-PrintUi '/o' $Printer
      Emit @('OK', 'show-queue', $Printer)
    }
    default { Fail ('unknown action: ' + $Action) 'unknownAction' }
  }
} catch {
  Emit @('ERR', (Get-ErrorText $_))
  exit 1
}

param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$Printer = '',
  [int]$Copies = 1,
  [string]$Mode = 'auto',
  [string]$OutputFile = '',
  [string]$TextEncoding = 'auto',
  [string]$Duplex = 'off',
  [string]$PageRange = 'all',
  [string]$Collate = 'true',
  [switch]$AllowDefaultSwitch,
  [int]$MaxPages = 200,
  [switch]$AllowHugeJob,
  [switch]$PlanOnly
)
# Prints a file with the Windows spooler.
#   -Path    file to print
#   -Printer target printer; empty = system default
#   -Copies  number of copies
#   -Mode    auto | gdi | queue | verb | file
#   -Duplex  off | long | short  (gdi: per-job request; queue/verb: changes the printer default
#            duplexing for the job, then restores it)
#   -PageRange  all | 1-3,5,8- | odd | even   (gdi only)
#   -Collate    true = collated copies (1,2,3,1,2,3), false = stacked (1,1,2,2,3,3)   (gdi only)
#   -MaxPages   page fuse: refuse a job above this many selected pages unless -AllowHugeJob is set
#   -AllowDefaultSwitch  allow temporarily changing the system default printer so programs that
#                        ignore the printer argument still print on the requested printer
#
# Four mechanisms, because no single one covers Windows:
#   gdi    - text and images are rendered here with System.Drawing and sent as a single spool job.
#            Only text and image files are accepted (binary content is refused) and a job above the
#            -MaxPages safety limit is refused unless -AllowHugeJob is passed.
#            This is the only path that can honour a page range, a copy order and a per-job duplex
#            request without touching the printer's default preferences.
#   queue  - plain text is sent to the spooler with Out-Printer, no GUI involved.
#   verb   - shell "printto" verb: the associated program receives the printer name. Programs that
#            silently ignore that argument (WPS Office does) are detected and refused instead of
#            pretending the job went to the requested printer.
#   file   - with -OutputFile the shell verb prints into a file (e.g. Microsoft Print to PDF).
#
# Records (US = 0x1f):
#   OK   <US> gdi <US> printer <US> file <US> pagesTotal <US> pagesSelected <US> copies <US> collate <US> duplex <US> observed <US> kind <US> outputFile
#   OK   <US> queue <US> printer <US> copies <US> file <US> duplex <US> observed
#   OK   <US> printto <US> printer <US> file <US> observed
#   OK   <US> defaultSwitch <US> printer <US> file <US> restored <US> observed <US> duplex
#   OK   <US> file <US> printer <US> outputFile <US> bytes <US> relocatedFrom <US> method
#   PLAN <US> text        (only with -PlanOnly; nothing is printed)
#   WARN <US> kind <US> text
#   CODE <US> code   (always immediately followed by ERR)
#   ERR  <US> message
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$US = [char]0x1f
function Emit([string[]]$fields) { [Console]::Out.WriteLine([string]::Join($US, $fields)) }
function Fail([string]$message, [string]$code) {
  if ($code) { Emit @('CODE', $code) }
  Emit @('ERR', $message)
  exit 1
}

# The GDI+ backend (page range, copy order, per-job duplex) lives in its own file, like scan.ps1
# and pdf.ps1: it is dot-sourced here, so its functions share this script's scope.
. (Join-Path $PSScriptRoot 'print-gdi.ps1')
# Registry-association probing, the system default printer and the printer status tables are shared
# with list-printers.ps1 and printer-action.ps1; they live in one file so the copies cannot drift.
. (Join-Path $PSScriptRoot 'print-common.ps1')

# A restore that has not run yet is written to disk before the printer is changed. A timeout kills
# this process (and its children), so the finally blocks below never get to run: without this file the
# printer's default duplex preference - and the system default printer - could stay changed for good
# (a real HP M283cdw was left on TwoSidedLongEdge that way). The next invocation restores it first.
function Get-RestoreStatePath { return (Join-Path $env:TEMP 'dsh-print-restore.json') }

function Save-RestoreState($state) {
  try { [System.IO.File]::WriteAllText((Get-RestoreStatePath), ($state | ConvertTo-Json -Compress)) } catch { }
}

function Clear-RestoreState {
  try { Remove-Item -LiteralPath (Get-RestoreStatePath) -Force -ErrorAction SilentlyContinue } catch { }
}

# Undoes whatever a killed run left behind. Returns a short comma-separated note for the caller.
function Restore-PendingState {
  $path = Get-RestoreStatePath
  if (-not (Test-Path -LiteralPath $path)) { return '' }
  $state = $null
  try { $state = (Get-Content -LiteralPath $path -Raw) | ConvertFrom-Json } catch { Clear-RestoreState; return '' }
  $notes = @()
  if ($state -and $state.printer) {
    $printer = [string]$state.printer
    if ($state.duplex -and ([string]$state.duplex) -ne 'off') {
      $info = @{ applied = $true; restore = [string]$state.duplex }
      if ((Restore-Duplex $printer $info) -eq 'ok') { $notes += 'duplex' } else { $notes += 'duplex-failed' }
    }
    if ($state.defaultPrinter) {
      $wanted = [string]$state.defaultPrinter
      if ((Get-DefaultPrinterName) -ne $wanted) {
        if (Set-DefaultPrinterRaw $wanted) { $notes += 'defaultPrinter' } else { $notes += 'defaultPrinter-failed' }
      }
    }
  }
  Clear-RestoreState
  return ($notes -join ',')
}

function Set-DefaultPrinterRaw([string]$printer) {
  try {
    if (-not ('DshPrintApi.Native' -as [type])) {
      Add-Type -Namespace 'DshPrintApi' -Name 'Native' -MemberDefinition '[DllImport("winspool.drv", CharSet=CharSet.Unicode, SetLastError=true)] public static extern bool SetDefaultPrinter(string pszPrinter);'
    }
    if ([DshPrintApi.Native]::SetDefaultPrinter($printer)) { return 'winspool' }
  } catch { }
  try {
    $net = New-Object -ComObject WScript.Network
    $net.SetDefaultPrinter($printer)
    return 'wscript'
  } catch { }
  try {
    # A printer name with spaces must reach printui as one argument; and the exit code alone proves
    # nothing, so the result is read back before it is reported as done.
    $quoted = '"' + $printer + '"'
    Start-Process -FilePath 'rundll32.exe' -ArgumentList @('printui.dll,PrintUIEntry', '/y', '/n', $quoted) -Wait -WindowStyle Hidden
    if ((Get-DefaultPrinterName) -eq $printer) { return 'rundll32' }
  } catch { }
  return ''
}

# Reads the spooler queue of one printer. readable=false means the queue could not be read at all,
# which is reported separately from "the queue is empty": the two mean very different things.
function Get-QueueSnapshot([string]$printer) {
  $snapshot = @{ readable = $false; jobs = @() }
  $raw = $null
  try {
    $raw = @(Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop)
    $snapshot.readable = $true
  } catch {
    try {
      $raw = @(Get-WmiObject -Class Win32_PrintJob -ErrorAction Stop)
      $snapshot.readable = $true
    } catch {
      $snapshot.readable = $false
      return $snapshot
    }
  }
  $jobs = @()
  foreach ($j in $raw) {
    $name = [string]$j.Name
    $idx = $name.LastIndexOf(',')
    $owner = $name
    if ($idx -gt 0) { $owner = $name.Substring(0, $idx) }
    if ($owner -eq $printer) { $jobs += $j }
  }
  $snapshot.jobs = $jobs
  return $snapshot
}

# Jobs can leave the queue in under a second, so the only honest answer is whether a NEW job was
# observed. 'submitted' = seen, 'unknown' = the queue stayed readable but nothing showed up in time,
# 'unreadable' = the queue could not be read (never claim success/empty in that case).
function Wait-ForQueueJob([string]$printer, [int]$timeoutMs) {
  $before = Get-QueueSnapshot $printer
  if (-not $before.readable) { return 'unreadable' }
  $beforeIds = @()
  foreach ($j in $before.jobs) { $beforeIds += [string]$j.JobId }
  $deadline = (Get-Date).AddMilliseconds($timeoutMs)
  while ((Get-Date) -lt $deadline) {
    $snap = Get-QueueSnapshot $printer
    if (-not $snap.readable) { return 'unreadable' }
    foreach ($j in $snap.jobs) {
      if ($beforeIds -notcontains [string]$j.JobId) { return 'submitted' }
    }
    Start-Sleep -Milliseconds 350
  }
  return 'unknown'
}

# A caught error must never come out as an empty string: an empty message silently turned a failed
# duplex change into "duplex off" with no warning at all (caught on a real HP M283cdw).
# Duplex lives in the printer's default printing preferences, so it is changed for this job only
# and restored afterwards. Returns @{ applied; mode; error; restore }.
function Set-Duplex([string]$printer, [string]$mode) {
  $info = @{ applied = $false; mode = ''; error = ''; restore = '' }
  if (-not $mode -or $mode -eq 'off') { return $info }
  $map = @{ 'long' = 'TwoSidedLongEdge'; 'short' = 'TwoSidedShortEdge' }
  $want = $map[$mode]
  if (-not $want) { $info.error = 'unknown duplex mode: ' + $mode; return $info }
  try {
    $cfg = Get-PrintConfiguration -PrinterName $printer -ErrorAction Stop
    $previous = [string]$cfg.DuplexingMode
    # $null = keeps the cmdlets from leaking objects into this function's output, which would turn the
    # returned hashtable into an array and make the caller read "duplex off" out of thin air.
    $null = Set-PrintConfiguration -PrinterName $printer -DuplexingMode $want -ErrorAction Stop
    $after = Get-PrintConfiguration -PrinterName $printer -ErrorAction Stop
    $actual = [string]$after.DuplexingMode
    if ($actual -ne $want) {
      $info.error = 'the driver kept the setting at ' + $actual + ' instead of ' + $want
      return $info
    }
    $info.applied = $true
    $info.mode = $actual
    if ($previous -and $previous -ne $want) { $info.restore = $previous }
    return $info
  } catch {
    $info.error = Get-ErrorText $_
    return $info
  }
}

function Restore-Duplex([string]$printer, $info) {
  if (-not $info -or -not $info.applied -or -not $info.restore) { return 'n/a' }
  try {
    $null = Set-PrintConfiguration -PrinterName $printer -DuplexingMode $info.restore -ErrorAction Stop
    $after = Get-PrintConfiguration -PrinterName $printer -ErrorAction Stop
    if ([string]$after.DuplexingMode -ne $info.restore) { return 'failed' }
    return 'ok'
  } catch { return 'failed' }
}

# Read-TextSmart now lives in print-gdi.ps1 (dot-sourced above), together with the renderer that
# calls it, so that the GDI backend can be exercised on its own.

function Test-IsText([string]$path) {
  $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
  $textExt = @('.txt', '.log', '.csv', '.md', '.ini', '.json', '.xml', '.htm', '.html', '.ps1', '.bat', '.cmd', '.js', '.mjs', '.ts', '.css', '.yml', '.yaml', '.sql')
  return ($textExt -contains $ext)
}

# Print-to-file drivers sometimes ignore the requested path and save under a generated name. Only
# the output directory is searched - not the desktop, Documents, Downloads and TEMP, where an
# unrelated PDF written in the last seconds would have been moved over the caller's file - and the
# requested path itself is never a candidate.
function Find-RecentPdf([datetime]$since, [string]$preferred) {
  if (-not $preferred) { return $null }
  $out = [System.IO.Path]::GetFullPath($preferred)
  $dir = [System.IO.Path]::GetDirectoryName($out)
  if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return $null }
  $newest = $null
  try {
    foreach ($candidate in @(Get-ChildItem -LiteralPath $dir -Filter '*.pdf' -File -ErrorAction SilentlyContinue)) {
      if ($candidate.FullName -eq $out) { continue }
      if ($candidate.Length -le 0) { continue }
      if ($candidate.LastWriteTime -lt $since) { continue }
      if (-not $newest -or $candidate.LastWriteTime -gt $newest.LastWriteTime) { $newest = $candidate }
    }
  } catch { }
  return $newest
}

try {
  $copies = [Math]::Max(1, $Copies)
  $target = ''
  if ($Printer) { $target = $Printer.Trim().Trim('"') }
  $full = [System.IO.Path]::GetFullPath($Path)
  if (-not (Test-Path -LiteralPath $full)) { Fail ('file not found: ' + $full) 'fileNotFound' }
  $item = Get-Item -LiteralPath $full
  if ($item.PSIsContainer) { Fail ('not a file: ' + $full) 'notAFile' }

  # NOTE: never name this local $mode. PowerShell variable names are case-insensitive, so it would
  # be the same slot as the [string]$Mode parameter, and "$mode = 'auto'" would overwrite whatever
  # the caller asked for - the mode argument was silently ignored that way until 1.2.0.
  $useMode = 'auto'
  if ($Mode) { $useMode = $Mode.ToLowerInvariant() }
  if (@('auto', 'gdi', 'queue', 'verb', 'file') -notcontains $useMode) { Fail ('unknown mode: ' + $Mode) 'unknownMode' }

  $duplexMode = 'off'
  if ($Duplex) { $duplexMode = $Duplex.ToLowerInvariant() }
  if (@('off', 'long', 'short') -notcontains $duplexMode) { Fail ('unknown duplex mode: ' + $Duplex) 'unknownDuplex' }

  # A run that was killed mid-job (see the restore-state helpers above) may have left the default
  # duplex preference or the system default printer changed. Undo that before anything else.
  $recovered = Restore-PendingState
  if ($recovered) { Emit @('WARN', 'restore', ('a previous print was interrupted; its printer settings were restored: ' + $recovered)) }

  # Page selection and copy order are only possible when the pages are rendered here (mode gdi).
  # 'all' and collated copies are the defaults, so only explicit requests count as demands.
  $pageSpec = ''
  if ($PageRange) { $pageSpec = $PageRange.Trim().ToLowerInvariant() }
  if ($pageSpec -eq 'all' -or $pageSpec -eq '*') { $pageSpec = '' }
  $collateText = 'true'
  if ($Collate -and $Collate.Trim().ToLowerInvariant() -eq 'false') { $collateText = 'false' }
  $stacked = ($collateText -eq 'false')
  $explicitSelection = ($pageSpec -ne '') -or ($stacked -and $copies -gt 1)

  $isText = Test-IsText $full
  $isImage = Test-IsImage $full
  $hasOutput = [bool]$OutputFile
  if ($useMode -eq 'auto') {
    if ($hasOutput) { if ($explicitSelection) { $useMode = 'gdi' } else { $useMode = 'file' } }
    elseif ($isText -or $isImage) { $useMode = 'gdi' }
    else { $useMode = 'verb' }
  }
  # The GDI path renders text and images only, and it used to decide "text" from the extension
  # alone, so an explicit "mode gdi" on a PDF slid past the auto-selection rule and the text branch
  # then decoded the PDF as text: a 15-page report reached the spooler as a 932-page job of binary
  # garbage (1.2.1 incident). print-gdi.ps1 sniffs the content as well; this is the first gate.
  if ($useMode -eq 'gdi' -and -not ($isText -or $isImage)) {
    Fail ('mode gdi renders text and image files only, and this file is neither. Print it with mode verb (when its associated program accepts a printer argument), or convert it to text or an image first: ' + $full) 'notText'
  }
  if ($useMode -eq 'file' -and -not $hasOutput) { Fail 'mode file needs an output path' 'outputRequired' }
  if ($explicitSelection -and @('queue', 'verb', 'file') -contains $useMode) {
    $what = 'a page range'
    if ($pageSpec -eq '') { $what = 'page-by-page copy order (collate=false)' }
    Fail ($what + ' cannot be done in mode ' + $useMode + ': that path hands the document to the shell verb (or sends text straight to the spooler) and gives the plugin no control over pages. Use mode gdi (text and image files), or print a text or image copy.') 'unsupportedSelection'
  }

  # -PlanOnly means "report what this job would do, print nothing". Only the GDI renderer can produce
  # that plan: on the other paths the flag used to be ignored and the job was printed for real, which
  # is exactly what the caller asked to avoid.
  if ($PlanOnly -and $useMode -ne 'gdi') {
    Fail ('planOnly needs mode gdi, because only the GDI renderer can report a page plan; mode ' + $useMode + ' cannot. Nothing was printed - pass mode gdi (text and image files) or drop planOnly.') 'planOnlyUnsupported'
  }

  if (-not $target) { $target = Get-DefaultPrinterName }
  if (-not $target) { Fail 'no default printer: pass a printer name explicitly' 'noPrinter' }

  $recentSince = (Get-Date).AddSeconds(-2)

  if ($useMode -eq 'gdi') {
    $plan = Invoke-GdiPrint -Path $full -Printer $target -PageRange $pageSpec -Copies $copies -Collate $collateText -Duplex $duplexMode -OutputFile $OutputFile -TextEncoding $TextEncoding -MaxPages $MaxPages -AllowHugeJob:$AllowHugeJob -PlanOnly:$PlanOnly
    if ($plan.ok) {
      if ($PlanOnly) {
        Emit @('PLAN', $plan.plan)
        exit 0
      }
      if ($plan.renamed) { Emit @('WARN', 'printer', ('this printer is registered as "' + $plan.printer + '" for GDI printing, and the job was sent there')) }
      $observed = Wait-ForQueueJob $target 2500
      Emit @('OK', 'gdi', $plan.printer, $full, [string]$plan.pagesTotal, $plan.pagesPrinted, [string]$copies, $plan.collate, $plan.duplex, $observed, $plan.kind, $plan.output)
      exit 0
    }
    # No silent downgrade. Falling back to the queue/verb path turned "the renderer refused this
    # file" into a real print through a path that cannot honour the page range or the copy order -
    # and, with -PlanOnly, into a real print as well. The failure is reported instead; the caller
    # can ask for the other mode explicitly.
    $hint = ''
    if ($isText) { $hint = ' Retry with mode=queue to send the plain text straight to the spooler.' }
    if ($isImage) { $hint = ' Retry with mode=verb to hand the image to its associated program.' }
    Fail ($plan.error + $hint) $plan.code
  }

  if ($useMode -eq 'queue') {
    # This path cannot select pages, so it gets the same content probe and page fuse as the renderer:
    # a mis-detected file must not reach the spooler as thousands of pages here either.
    $probe = Get-TextProbe $full
    if ($probe.kind -eq 'unreadable') { Fail ('the file could not be read: ' + $probe.error) 'fileUnreadable' }
    if ($probe.kind -eq 'binary') {
      Fail 'this file is not text: it starts with a binary signature or contains NUL bytes without a text byte-order mark, so it cannot be sent through the spooler text path. Convert it to text first.' 'notText'
    }
    $text = Read-TextSmart $full $TextEncoding
    $estimate = 1
    $lineCount = @($text -split "\n").Count
    if ($lineCount -gt 60) { $estimate = [int][Math]::Ceiling($lineCount / 60.0) }
    if (-not $AllowHugeJob -and $estimate -gt $MaxPages) {
      Fail ('this text would print about ' + $estimate + ' pages, far more than expected: the ' + $MaxPages + '-page safety limit keeps a mis-detected file from emptying a paper tray. Check the file, then pass allowHugeJob to print it anyway.') 'tooManyPages'
    }
    $temp = Join-Path $env:TEMP ('dsh-print-' + [System.Guid]::NewGuid().ToString('N') + '.txt')
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($temp, $text, $utf8)
    # NOTE: never name this local $duplex. PowerShell variable names are case-insensitive, so it
    # would be the same slot as the [string]$Duplex parameter: assigning a hashtable to a typed
    # string variable silently converts it to the literal text "System.Collections.Hashtable",
    # every $duplex.applied lookup then yields $null (reported as duplex=off) and the restore step
    # is skipped, leaving the printer's default duplex preference changed for good.
    $duplexInfo = Set-Duplex $target $duplexMode
    if ($duplexInfo.restore) { Save-RestoreState @{ printer = $target; duplex = $duplexInfo.restore } }
    try {
      for ($i = 0; $i -lt $copies; $i++) {
        Get-Content -LiteralPath $temp -Encoding UTF8 | Out-Printer -Name $target
      }
    } finally {
      Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
      $restore = Restore-Duplex $target $duplexInfo
      if ($restore -eq 'ok' -or -not $duplexInfo.restore) { Clear-RestoreState }
    }
    if ($duplexInfo.error) { Emit @('WARN', 'duplex', ('could not set duplex on ' + $target + ': ' + $duplexInfo.error)) }
    if ($duplexInfo.applied -and $restore -eq 'failed') { Emit @('WARN', 'duplex', 'duplex was set but could not be restored to the previous value') }
    $observed = Wait-ForQueueJob $target 2500
    $duplexText = 'off'
    if ($duplexInfo.applied) { $duplexText = $duplexInfo.mode }
    Emit @('OK', 'queue', $target, [string]$copies, $full, $duplexText, $observed)
    exit 0
  }

  if ($useMode -eq 'file') {
    $out = [System.IO.Path]::GetFullPath($OutputFile)
    $outDir = [System.IO.Path]::GetDirectoryName($out)
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
    # A shell verb receives one command line, not an argv array, so the printer name and the output
    # path are quoted here with their embedded quotes removed - caller text never reaches a command
    # line raw.
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $full
    $psi.Verb = 'PrintTo'
    $psi.UseShellExecute = $true
    $psi.Arguments = '"' + ($target -replace '"', '') + '" "' + ($out -replace '"', '') + '"'
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.WaitForExit()
    $size = 0
    for ($i = 0; $i -lt 40; $i++) {
      Start-Sleep -Milliseconds 500
      if (Test-Path -LiteralPath $out) {
        $len = (Get-Item -LiteralPath $out).Length
        if ($len -gt 0 -and $len -eq $size) { break }
        $size = $len
      }
    }
    $relocated = ''
    if (-not (Test-Path -LiteralPath $out) -or (Get-Item -LiteralPath $out).Length -le 0) {
      $found = Find-RecentPdf $recentSince $out
      if (-not $found) { Fail ('the printer did not write ' + $out + ': the associated program ignored the print-to-file request') 'noOutput' }
      if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
      Move-Item -LiteralPath $found.FullName -Destination $out -Force
      $relocated = $found.FullName
    }
    $size = (Get-Item -LiteralPath $out).Length
    Emit @('OK', 'file', $target, $out, [string]$size, $relocated, 'printto')
    exit 0
  }

  # mode verb: check that the shell verb will actually honour the printer name before running it.
  $verbInfo = Get-PrintVerb $full
  $verb = $verbInfo.verb
  $ignored = Test-PrinterArgIgnored $verbInfo.command
  if ($verb -eq 'printto' -and -not $ignored) {
    # NOTE: never name this local $duplex. PowerShell variable names are case-insensitive, so it
    # would be the same slot as the [string]$Duplex parameter: assigning a hashtable to a typed
    # string variable silently converts it to the literal text "System.Collections.Hashtable",
    # every $duplex.applied lookup then yields $null (reported as duplex=off) and the restore step
    # is skipped, leaving the printer's default duplex preference changed for good.
    $duplexInfo = Set-Duplex $target $duplexMode
    if ($duplexInfo.restore) { Save-RestoreState @{ printer = $target; duplex = $duplexInfo.restore } }
    try {
      $psi = New-Object System.Diagnostics.ProcessStartInfo
      $psi.FileName = $full
      $psi.Verb = 'PrintTo'
      $psi.UseShellExecute = $true
      $psi.Arguments = '"' + ($target -replace '"', '') + '"'
      $proc = [System.Diagnostics.Process]::Start($psi)
      $proc.WaitForExit()
    } finally {
      $restore = Restore-Duplex $target $duplexInfo
      if ($restore -eq 'ok' -or -not $duplexInfo.restore) { Clear-RestoreState }
    }
    if ($duplexInfo.error) { Emit @('WARN', 'duplex', ('could not set duplex on ' + $target + ': ' + $duplexInfo.error)) }
    $observed = Wait-ForQueueJob $target 2500
    Emit @('OK', 'printto', $target, $full, $observed)
    exit 0
  }

  if (-not $AllowDefaultSwitch) {
    $reason = 'this file type has no printto verb, so the shell would print to the default printer'
    if ($ignored) { $reason = 'the associated program (' + $verbInfo.progId + ') ignores the printer argument, so this would print to the default printer instead' }
    Fail ($reason + '. Pass allowDefaultSwitch to switch the system default printer for this job, or print a text/pdf copy instead.') 'noPrintToVerb'
  }

  # Escape hatch: make the requested printer the system default, print, put the old one back.
  $previous = Get-DefaultPrinterName
  $switched = ''
  if ($previous -and $previous -ne $target) {
    # The old default is written to the restore file BEFORE the switch: if this process is killed in
    # between, the next run still knows what to put back.
    Save-RestoreState @{ printer = $target; defaultPrinter = $previous }
    $switched = Set-DefaultPrinterRaw $target
    if (-not $switched) { Fail ('could not make ' + $target + ' the default printer') 'switchFailed' }
  }
  $duplexInfo = Set-Duplex $target $duplexMode
  $pending = @{ printer = $target }
  if ($duplexInfo.restore) { $pending.duplex = $duplexInfo.restore }
  if ($previous -and $previous -ne $target) { $pending.defaultPrinter = $previous }
  Save-RestoreState $pending
  $restored = 'n/a'
  try {
    Start-Process -FilePath $full -Verb Print -Wait
  } finally {
    $duplexBack = Restore-Duplex $target $duplexInfo
    if ($previous -and $previous -ne $target) {
      $back = Set-DefaultPrinterRaw $previous
      if ($back -and (Get-DefaultPrinterName) -eq $previous) { $restored = 'ok' } else { $restored = 'failed' }
    }
    if ($restored -ne 'failed' -and $duplexBack -ne 'failed') { Clear-RestoreState }
  }
  if ($duplexInfo.applied -and $duplexBack -eq 'failed') { Emit @('WARN', 'duplex', 'duplex was set but could not be restored to the previous value') }
  if ($duplexInfo.error) { Emit @('WARN', 'duplex', ('could not set duplex on ' + $target + ': ' + $duplexInfo.error)) }
  if ($restored -eq 'failed') { Emit @('WARN', 'default', ('the system default printer could not be set back to ' + $previous)) }
  $observed = Wait-ForQueueJob $target 2500
  $duplexText = 'off'
  if ($duplexInfo.applied) { $duplexText = $duplexInfo.mode }
  Emit @('OK', 'defaultSwitch', $target, $full, $restored, $observed, $duplexText)
  exit 0
} catch {
  Emit @('CODE', 'printFailed')
  Emit @('ERR', (Error-Text $_))
  exit 1
}

# Regression checks for the 1.3.x safety guards. Uses no spooler and prints nothing.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$psDir = Join-Path (Split-Path -Parent $here) 'ps'
. (Join-Path $psDir 'print-gdi.ps1')

$work = Join-Path $env:TEMP ('dsh-guard-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -Force | Out-Null

function Write-Bytes([string]$path, [int[]]$head, [int]$tail) {
  $list = New-Object System.Collections.Generic.List[byte]
  foreach ($b in $head) { $list.Add([byte]$b) }
  for ($i = 0; $i -lt $tail; $i++) { $list.Add([byte]($i % 251)) }
  [System.IO.File]::WriteAllBytes($path, $list.ToArray())
}

$fakePdf = Join-Path $work 'fake.pdf'
Write-Bytes $fakePdf @(0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34, 0x0A) 400
$fakeDocx = Join-Path $work 'fake.docx'
Write-Bytes $fakeDocx @(0x50, 0x4B, 0x03, 0x04, 0x14, 0x00) 200
$fakeOle = Join-Path $work 'fake.xls'
Write-Bytes $fakeOle @(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1) 120
$nulText = Join-Path $work 'binary-nul.txt'
[System.IO.File]::WriteAllBytes($nulText, [byte[]](0x41, 0x42, 0x00, 0x43, 0x44))
$twoPages = Join-Path $work 'two-pages.txt'
[System.IO.File]::WriteAllText($twoPages, ('ALPHA' + [char]10 + [char]12 + 'BETA' + [char]10), (New-Object System.Text.UTF8Encoding $false))
$plain = Join-Path $work 'plain.txt'
[System.IO.File]::WriteAllText($plain, 'hello world' + [char]10, (New-Object System.Text.UTF8Encoding $false))

$script:fails = 0
function Check([string]$label, $actual, $expect, $detail) {
  $ok = ($actual -eq $expect)
  if (-not $ok) { $script:fails++ }
  $verdict = 'PASS'
  if (-not $ok) { $verdict = 'FAIL' }
  $extra = ''
  if (-not $ok -and $detail) { $extra = ' [' + [string]$detail + ']' }
  Write-Host ($verdict + ' ' + $label + ' -> ' + $actual + ' (want ' + $expect + ')' + $extra)
}

# A nested powershell.exe started as "& powershell.exe ... 2>&1" can come back with no output at
# all in restricted sandboxes (the child's streams are pipes). The child therefore writes to files
# that are read back afterwards, which works everywhere.
function Invoke-PrintScript([string[]]$scriptArgs) {
  $stdout = Join-Path $work ('child-' + [guid]::NewGuid().ToString('N') + '.out')
  $stderr = $stdout + '.err'
  $exe = Join-Path $PSHOME 'powershell.exe'
  $childArgs = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File') + $scriptArgs
  # Start-Process joins the array with spaces and does not quote it, so an argument that contains
  # spaces (printer names such as "Microsoft XPS Document Writer") has to be quoted here.
  $quoted = @()
  foreach ($arg in $childArgs) {
    if ([string]$arg -match '[\s]') { $quoted += ('"' + [string]$arg + '"') } else { $quoted += [string]$arg }
  }
  Start-Process -FilePath $exe -ArgumentList ($quoted -join ' ') -Wait -NoNewWindow -RedirectStandardOutput $stdout -RedirectStandardError $stderr | Out-Null
  $lines = @()
  foreach ($file in @($stdout, $stderr)) {
    if (Test-Path -LiteralPath $file) {
      $lines += [System.IO.File]::ReadAllLines($file)
      Remove-Item -LiteralPath $file -Force
    }
  }
  return $lines
}

$target = 'Microsoft XPS Document Writer'
$r = Invoke-GdiPrint -Path $fakePdf -Printer $target -PlanOnly
Check 'gdi refuses pdf' $r.code 'notText' $r.error
$r = Invoke-GdiPrint -Path $fakeDocx -Printer $target -PlanOnly
Check 'gdi refuses docx' $r.code 'notText' $r.error
$r = Invoke-GdiPrint -Path $fakeOle -Printer $target -PlanOnly
Check 'gdi refuses ole' $r.code 'notText' $r.error
$r = Invoke-GdiPrint -Path $nulText -Printer $target -PlanOnly
Check 'gdi refuses text with NUL bytes' $r.code 'notText' $r.error
$r = Invoke-GdiPrint -Path $twoPages -Printer $target -PlanOnly -MaxPages 1
Check 'page fuse trips' $r.code 'tooManyPages' $r.error
$r = Invoke-GdiPrint -Path $twoPages -Printer $target -PlanOnly -MaxPages 1 -AllowHugeJob
Check 'page fuse override' $r.ok $true ($r.code + ' ' + $r.error)
$r = Invoke-GdiPrint -Path $twoPages -Printer $target -PlanOnly
Check 'two-page text still plans' $r.ok $true ($r.code + ' ' + $r.error)
Check 'two-page text is 2 pages' $r.pagesTotal 2 $r.error
$r = Invoke-GdiPrint -Path $plain -Printer $target -PlanOnly
Check 'plain text still plans' $r.ok $true ($r.code + ' ' + $r.error)

$scriptPath = Join-Path $psDir 'print-file.ps1'
$out = Invoke-PrintScript @($scriptPath, '-Path', $fakePdf, '-Mode', 'gdi', '-PlanOnly')
$codeLine = ($out | Where-Object { [string]$_ -like 'CODE*' } | Select-Object -First 1)
Check 'print-file gate 1 refuses pdf' ([bool]([string]$codeLine -like '*notText*')) $true ([string]($out -join ' | '))
$out = Invoke-PrintScript @($scriptPath, '-Path', $twoPages, '-Mode', 'gdi', '-Printer', $target, '-PlanOnly')
$planLine = ($out | Where-Object { [string]$_ -like 'PLAN*' } | Select-Object -First 1)
Check 'print-file gdi still plans text' ([bool]($planLine -like 'PLAN*')) $true ([string]($out -join ' | '))

# ---- 1.3.0 additions ---------------------------------------------------
# The registry verb lookup is the code path the review found completely broken (a missing
# backslash made every lookup fail, so WPS documents were always rejected). If this machine has
# no association at all the lookup legitimately returns nothing, so report SKIP instead of FAIL.
. (Join-Path $psDir 'print-common.ps1')
$verbPdf = Get-PrintVerb $fakePdf
if ($verbPdf -and $verbPdf.verb) {
  Check 'registry print verb for .pdf' ([bool]$verbPdf.verb) $true $verbPdf.command
} else {
  Write-Host 'SKIP registry print verb for .pdf (no association on this machine)'
}
$verbDocx = Get-PrintVerb $fakeDocx
if ($verbDocx -and $verbDocx.verb) {
  Check 'registry print verb for .docx' ([bool]$verbDocx.verb) $true $verbDocx.command
} else {
  Write-Host 'SKIP registry print verb for .docx (no association on this machine)'
}

# UTF-16 text is full of NUL bytes; the old NUL rule refused it, which made encoding=unicode
# unreachable. The BOM must win.
$utf16 = Join-Path $work 'utf16.txt'
[System.IO.File]::WriteAllText($utf16, 'hello world', (New-Object System.Text.UnicodeEncoding $false, $true))
$probe = Get-TextProbe $utf16
Check 'utf-16le text is not binary' $probe.kind 'text' $probe.error
Check 'utf-16le BOM recognised' $probe.bom 'utf16le' ''

# The extension must not be trusted: a PDF renamed to .txt is still binary.
$renamed = Join-Path $work 'actually-pdf.txt'
Copy-Item -LiteralPath $fakePdf -Destination $renamed -Force
$r = Invoke-GdiPrint -Path $renamed -Printer $target -PlanOnly
Check 'pdf renamed to .txt is refused' $r.code 'notText' $r.error

# -PlanOnly used to be honoured only on the gdi path; the queue path silently printed. It must
# now be refused up front.
$out2 = Invoke-PrintScript @($scriptPath, '-Path', $plain, '-Mode', 'queue', '-PlanOnly')
$codeLine2 = ($out2 | Where-Object { [string]$_ -like 'CODE*' } | Select-Object -First 1)
Check 'planOnly refuses queue path' ([bool]([string]$codeLine2 -like '*planOnlyUnsupported*')) $true ([string]($out2 -join ' | '))

# Cancel and purge must report what a read-back saw, never a bare "deleted" claim.
$jobsSrc = [System.IO.File]::ReadAllText((Join-Path $psDir 'printer-jobs.ps1'))
Check 'cancel verifies by read-back' ([bool]($jobsSrc -match 'stillQueued')) $true
Check 'purge reports remaining' ([bool]($jobsSrc -match 'remaining')) $true
$fileSrc = [System.IO.File]::ReadAllText($scriptPath)
Check 'print-file exposes MaxPages' ([bool]($fileSrc -match 'MaxPages')) $true

# ---- 1.3.1 additions ---------------------------------------------------
# The unknown-printer text used to run straight into the suggestion ("... Fax Retry with
# mode=queue ..."), because the available list had no closing period.
$r = Invoke-GdiPrint -Path $plain -Printer 'NoSuchPrinter-For-Guard' -PlanOnly
Check 'unknown printer is an error' $r.code 'unknownPrinter' $r.error
Check 'unknown printer message ends with a period' ([bool]([string]$r.error -match '\.$')) $true $r.error

# printer_info used to read Win32_Printer.Duplex / .Color only; a driver that does not publish
# those properties looked like "unknown capabilities". Both fields must now say where the value
# came from so the host can label it as inferred.
$psExe = Join-Path $PSHOME 'powershell.exe'
$capsOut = & $psExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $psDir 'printer-caps.ps1') -Printer $target 2>&1
$capsLine = ($capsOut | Where-Object { [string]$_ -like ('CAPS' + [char]0x1f + '*') } | Select-Object -First 1)
if ($capsLine) {
  $capsFields = ([string]$capsLine).Split([char]0x1f)
  Check 'caps record carries 12 fields' $capsFields.Count 12 ([string]$capsLine)
  Check 'caps says where duplex came from' ([bool]($capsFields[10])) $true ([string]$capsLine)
  Check 'caps says where colour came from' ([bool]($capsFields[11])) $true ([string]$capsLine)
} else {
  Write-Host 'SKIP printer capability sources (printer not visible to WMI on this machine)'
}
$capsSrc = [System.IO.File]::ReadAllText((Join-Path $psDir 'printer-caps.ps1'))
Check 'caps falls back to capability bits' ([bool]($capsSrc -match 'capabilities')) $true
Check 'caps falls back to the gdi driver' ([bool]($capsSrc -match 'PrinterSettings')) $true
$printersSrc = [System.IO.File]::ReadAllText((Join-Path $psDir '..\lib\printers.js'))
Check 'host parses capability sources' ([bool]($printersSrc -match 'duplexSource = caps\[10\]')) $true

# A stale eSCL entry must not make every default scan fail: the list marks each device and the
# scan picks the first one that opens.
$scanOut = & $psExe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $psDir 'list-scanners.ps1') 2>&1
$scanLines = @($scanOut | Where-Object { [string]$_ -like ('SCANNER' + [char]0x1f + '*') })
if ($scanLines.Count -gt 0) {
  $badScan = @($scanLines | Where-Object { ([string]$_).Split([char]0x1f).Count -ne 6 })
  Check 'scanner records carry availability' $badScan.Count 0 ([string]($scanLines -join ' | '))
} else {
  Write-Host 'SKIP scanner availability (no scanner or WIA unavailable on this machine)'
}
$scanSrc = [System.IO.File]::ReadAllText((Join-Path $psDir 'scan.ps1'))
Check 'scan skips devices that will not open' ([bool]($scanSrc -match 'no scanner could be opened')) $true

# ---- 1.3.2 additions ---------------------------------------------------
# The HP PCL-6 driver stamps 96 dpi into every page header while really scanning at 300 dpi. Both
# the saved page and the A4 crop has to use the negotiated resolution, otherwise every scan is
# mislabelled and crop=a4 shrinks an A4 page to a quarter of its size.
. (Join-Path $psDir 'image-common.ps1')
Add-Type -AssemblyName System.Drawing
$dpiBmp = Join-Path $work 'dpi-test.bmp'
$b1 = New-Object System.Drawing.Bitmap -ArgumentList @(1200, 1700)
try { $b1.SetResolution(96, 96); $b1.Save($dpiBmp, [System.Drawing.Imaging.ImageFormat]::Bmp) } finally { $b1.Dispose() }
$dpiJpg = Join-Path $work 'dpi-test.jpg'
Copy-Item -LiteralPath $dpiBmp -Destination $dpiJpg
Convert-ImageFile $dpiJpg 'jpg' 300 | Out-Null
$dpiSize = Get-ImageSize $dpiJpg
Check 'conversion writes the negotiated dpi' $dpiSize.dpi 300

$cropJpg = Join-Path $work 'crop-a4.jpg'
$b2 = New-Object System.Drawing.Bitmap -ArgumentList @(2479, 3507)
try { $b2.SetResolution(96, 96); $b2.Save($cropJpg, [System.Drawing.Imaging.ImageFormat]::Jpeg) } finally { $b2.Dispose() }
$cropRes = Crop-Image $cropJpg 300 'a4' 'jpg'
$cropSize = Get-ImageSize $cropJpg
Check 'a4 crop uses the negotiated dpi' $cropSize.width 2479 ([string]$cropRes.dpi)
Check 'a4 crop leaves an a4 page alone' $cropRes.cropped $false

$longJpg = Join-Path $work 'crop-long.jpg'
$b3 = New-Object System.Drawing.Bitmap -ArgumentList @(2479, 4200)
try { $b3.SetResolution(300, 300); $b3.Save($longJpg, [System.Drawing.Imaging.ImageFormat]::Jpeg) } finally { $b3.Dispose() }
$longRes = Crop-Image $longJpg 300 'a4' 'jpg'
$longSize = Get-ImageSize $longJpg
Check 'a4 crop trims a long feeder page' $longSize.height 3508 ([string]$longRes.dpi)
Check 'a4 crop reports the trim' $longRes.cropped $true

# ---- 1.3.3 additions ---------------------------------------------------
# The feeder used to scan exactly one sheet: the driver defaults Pages=1, and every page is a fresh
# WIA item whose properties go back to the driver defaults. The A4 extent and the prefix match are
# pure functions, so they can be pinned here without a scanner.
. (Join-Path $psDir 'wia-common.ps1')
$a4 = Get-A4Extent 300
Check 'a4 extent at 300 dpi' ([string]$a4.width + 'x' + [string]$a4.height) '2480x3508'
$fakeDevices = @(
  @{ index = 1; name = 'Alpha Scanner'; id = 'id-alpha'; type = 1 },
  @{ index = 2; name = 'Beta Scanner'; id = 'id-beta'; type = 1 }
)
$pick = Find-WiaDevice $fakeDevices 'Alpha'
Check 'scanner prefix match picks one' ([bool]($pick.device -and $pick.device.name -eq 'Alpha Scanner' -and -not $pick.ambiguous)) $true ([string]($pick.names -join ' | '))
$fakeTwins = @(
  @{ index = 1; name = 'Scanner One'; id = 'id-one'; type = 1 },
  @{ index = 2; name = 'Scanner Two'; id = 'id-two'; type = 1 }
)
$amb = Find-WiaDevice $fakeTwins 'Scanner'
Check 'ambiguous scanner prefix is reported' ([bool]$amb.ambiguous) $true ([string]($amb.names -join ' | '))

# ---- 1.3.4 additions ---------------------------------------------------
# The feeder used to be told DPS_PAGES = 0 ("keep feeding until the tray is empty"). This HP WSD
# driver answers every transfer after that with E_INVALIDARG, so multi-page scanning now opens one
# WIA session per sheet, never writes the page count, and stops on the PAPER_EMPTY status instead.
$scanSrc34 = [System.IO.File]::ReadAllText((Join-Path $psDir 'scan.ps1'))
$wiaSrc34 = [System.IO.File]::ReadAllText((Join-Path $psDir 'wia-common.ps1'))
Check 'scan never zeroes the feeder page count' ([bool]($scanSrc34 -match 'Set-WiaFeederAllPages')) $false
Check 'wia helper no longer writes 3096' ([bool]($wiaSrc34 -match 'Properties 3096')) $false
Check 'scan opens a session per sheet' ([bool](($scanSrc34 -match 'function Open-ScanItem') -and ($scanSrc34 -match 'Close-ScanItem'))) $true

# ---- 1.3.5 additions ---------------------------------------------------
# The feeder is now driven through eSCL instead of WIA, because this HP driver refuses multi-page
# WIA jobs. The pure parts (settings document, resolution pick, address handling) are pinned here;
# the network calls are only exercised by the hardware tests.
. (Join-Path $psDir 'escl-common.ps1')
$scanSrc35 = [System.IO.File]::ReadAllText((Join-Path $psDir 'scan.ps1'))
Check 'scan loads the eSCL client' ([bool]($scanSrc35 -match 'escl-common[.]ps1')) $true
Check 'scan implements an eSCL backend' ([bool](($scanSrc35 -match 'function Invoke-EsclScanPages') -and ($scanSrc35 -match 'function Finish-PageFile'))) $true
Check 'escl feeder token is Feeder' ([bool]($scanSrc35 -match "[=] 'Feeder'")) $true
Check 'escl platen token is Platen' ([bool]($scanSrc35 -match "[=] 'Platen'")) $true
$addr = Resolve-EsclBase ''
Check 'no address means no eSCL guess' $addr.ok $false
$xmlAdf = New-EsclSettingsXml -InputSource 'Feeder' -Resolution 300 -ColorMode 'RGB24'
Check 'escl settings ask for a4' ([bool](($xmlAdf -match '<pwg:Width>2480</pwg:Width>') -and ($xmlAdf -match '<pwg:Height>3508</pwg:Height>') -and ($xmlAdf -match 'escl:ThreeHundredthsOfInches'))) $true
Check 'escl settings carry the source' ([bool]($xmlAdf -match '<pwg:InputSource>Feeder</pwg:InputSource>')) $true
Check 'escl settings carry the resolution' ([bool](($xmlAdf -match '<scan:XResolution>300</scan:XResolution>') -and ($xmlAdf -match '<scan:YResolution>300</scan:YResolution>'))) $true
Check 'escl settings stay well formed' ([bool](($xmlAdf.StartsWith('<?xml')) -and ($xmlAdf.EndsWith('</scan:ScanSettings>')))) $true
Check 'escl rounds to a supported resolution' (Select-EsclResolution -Accepted @(75, 150, 300) -Want 1000) 300
Check 'escl keeps an exact resolution' (Select-EsclResolution -Accepted @(75, 150, 200, 300, 600, 1200) -Want 600) 600

Remove-Item -LiteralPath $work -Recurse -Force
if ($script:fails -gt 0) { Write-Host ('FAILURES: ' + $script:fails); exit 1 }
Write-Host 'ALL GUARD CHECKS PASSED'
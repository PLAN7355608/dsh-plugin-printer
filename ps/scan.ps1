param(
  [string]$Device = '',
  [Parameter(Mandatory = $true)][string]$Output,
  [int]$Dpi = 300,
  [string]$Mode = 'color',
  [string]$Format = 'jpg',
  [string]$Source = 'auto',
  [string]$Crop = 'off',
  [switch]$KeepPages
)
# Scans with a WIA device (flatbed or automatic document feeder).
#
# Two backends, picked automatically for the selected device:
#   * eSCL (network): a whole ADF stack runs as ONE job, so a multi-page feeder scan works; this is
#     the preferred path for the HP WSD/eSCL all-in-one this plugin was verified on.
#   * WIA (USB/legacy): one page per session; the ADF stops after one sheet on drivers that only
#     accept WIA_DPS_PAGES = 1, which is what the older builds were limited by.
#   -Device    device name, DeviceID, or 1-based index; empty = first scanner
#   -Output    target path: image series (page 2+ get -2, -3, ...) or the .pdf file
#   -Dpi       requested resolution; the driver decides what it really supports
#   -Mode      color | gray | text
#   -Format    jpg | png | tif | bmp | pdf  (pdf = one multi-page file, built by Write-Pdf)
#   -Source    auto | flat | feed
#   -Crop      off | a4  (trim the right/bottom overflow when the sheet is longer than A4)
#   -KeepPages keep the per-page image files that went into the pdf
#
# A multi-page pdf is only produced when the whole scan succeeded: pages are collected in a
# temporary directory and moved to their final names (or folded into the pdf) at the end.
# Pages that were finished before a failure are moved out anyway and reported as PAGE records, so a
# half-finished scan still leaves usable pages behind; a STAT record says how many.
#
# WIA scripting notes learned from real hardware:
#   * Properties.Item(n) takes a 1-based INDEX, not a property id, so lookups scan the list.
#   * Drivers ignore the requested format: Transfer() may hand back BMP, and some formats make
#     Transfer() fail outright ("Unspecified error"). Every page is inspected after the transfer.
#   * XRES/YRES only accept the values the device really supports, so the accepted set is probed.
#   * A document feeder that runs out of paper reports WIA_ERROR_PAPER_EMPTY (0x80210003) from
#     Transfer(). That is normal end-of-document, NOT a failure: everything scanned so far stays.
#   * DPS_DOCUMENT_HANDLING_SELECT (3088) holds whatever the driver used last, often FLATBED even
#     when the feeder is loaded, so -Source auto must not trust it.
#
# Output records (US = 0x1f):
#   PAGE <US> index <US> path <US> bytes <US> format <US> width <US> height <US> dpi <US> converted <US> kept
#   PDF  <US> path <US> bytes <US> pages <US> dpi
#   WARN <US> kind <US> text
#   STAT <US> pages <US> stoppedBy <US> device <US> source   (only before a failure, never on success)
#   OK   <US> scan <US> device <US> pages <US> requestedDpi <US> effectiveDpi <US> source <US> formats <US> resolutions <US> stoppedBy <US> caps <US> nativeFormat <US> crop <US> fallback
#   CODE <US> code   (machine readable reason, always immediately followed by ERR)
#   ERR  <US> message
#
# Error codes: unknownFormat, unknownSource, unknownCrop, unknownMode, scannerUnreadable,
# noScanner, scannerNotFound, noFeeder, noFlatbed, feederEmpty, noPage, transferFailed,
# saveFailed, pdfFailed, deviceError
. (Join-Path $PSScriptRoot 'pdf.ps1')
. (Join-Path $PSScriptRoot 'image-common.ps1')
. (Join-Path $PSScriptRoot 'wia-common.ps1')
. (Join-Path $PSScriptRoot 'escl-common.ps1')
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$US = [char]0x1f
function Emit([string[]]$fields) { [Console]::Out.WriteLine([string]::Join($US, $fields)) }
function Fail([string]$message, [string]$code) {
  if ($code) { Emit @('CODE', $code) }
  Emit @('ERR', $message)
  exit 1
}

# Upper bound on pages per call: a stuck feeder or a driver that keeps handing back the same sheet
# would otherwise scan until the disk is full.
$MAX_PAGES = 50

# WIA_ERROR_PAPER_EMPTY = 0x80210003: the feeder is empty or ran out of paper.
$HR_PAPER_EMPTY = -2145320957

function Get-ErrorCode($err) {
  if (-not $err) { return 0 }
  $code = 0
  try { $code = [int]$err.Exception.HResult } catch { $code = 0 }
  if ($code -eq 0) {
    try { if ($err.Exception.InnerException) { $code = [int]$err.Exception.InnerException.HResult } } catch { }
  }
  return $code
}

function Test-PaperEmpty($err) {
  if (-not $err) { return $false }
  if ((Get-ErrorCode $err) -eq $HR_PAPER_EMPTY) { return $true }
  $msg = ''
  try { $msg = [string]$err.Exception.Message } catch { $msg = '' }
  if ($msg -match 'no paper|paper empty|out of paper') { return $true }
  return $false
}

function Get-ErrorText($err) {
  if (-not $err) { return '' }
  $msg = ''
  try { $msg = [string]$err.Exception.Message } catch { $msg = '' }
  $code = Get-ErrorCode $err
  if ($code -ne 0) { return ($msg + ' (0x' + ([Convert]::ToString(($code -band 0xFFFFFFFFL), 16)) + ')') }
  return $msg
}


# Picks the supported resolution closest to the request and applies it.
function Select-Resolution($item, [int]$want) {
  $accepted = @(Get-AcceptedResolutions $item)
  if ($accepted.Count -eq 0) { return @{ effective = 0; accepted = @() } }
  $best = $accepted[0]
  $bestDelta = [Math]::Abs($accepted[0] - $want)
  foreach ($v in $accepted) {
    $delta = [Math]::Abs($v - $want)
    if ($delta -lt $bestDelta) { $best = $v; $bestDelta = $delta }
  }
  Set-PropertyById $item.Properties 6147 $best | Out-Null
  Set-PropertyById $item.Properties 6148 $best | Out-Null
  return @{ effective = $best; accepted = $accepted }
}

function Set-Source($dev, [string]$src) {
  if ($src -eq 'feed') { Set-PropertyById $dev.Properties 3088 1 | Out-Null }
  elseif ($src -eq 'flat') { Set-PropertyById $dev.Properties 3088 2 | Out-Null }
}

function Page-Path([string]$base, [int]$page) {
  if ($page -le 1) { return $base }
  $dir = [System.IO.Path]::GetDirectoryName($base)
  $stem = [System.IO.Path]::GetFileNameWithoutExtension($base)
  $ext = [System.IO.Path]::GetExtension($base)
  $name = $stem + '-' + [string]$page + $ext
  if ($dir) { return (Join-Path $dir $name) }
  return $name
}

# Moves the collected page files out of the temporary directory and emits their PAGE records.
#   pdfMode + keepPages = 0 : the pages went into the pdf, so only an empty path is reported and the
#                             files are left for the caller's cleanup.
#   partial = 1             : the pdf was never produced, so everything is kept as an image series.
function Move-ScanPages {
  param($state, [string]$imgBase, [bool]$pdfMode, [bool]$keepPages, [bool]$partial)
  foreach ($p in $state.pages) {
    if ($pdfMode -and -not $keepPages -and -not $partial) {
      Emit @('PAGE', [string]$p.page, '', [string]$p.bytes, $p.format, [string]$p.width, [string]$p.height, [string]$p.dpi, [string]$p.converted, '0')
      continue
    }
    $final = Page-Path $imgBase $p.page
    $final = [System.IO.Path]::ChangeExtension($final, '.' + $p.format)
    if (Test-Path -LiteralPath $final) { Remove-Item -LiteralPath $final -Force }
    Move-Item -LiteralPath $p.temp -Destination $final -Force
    $bytes = 0
    if (Test-Path -LiteralPath $final) { $bytes = (Get-Item -LiteralPath $final).Length }
    Emit @('PAGE', [string]$p.page, $final, [string]$bytes, $p.format, [string]$p.width, [string]$p.height, [string]$p.dpi, [string]$p.converted, '1')
  }
}

# Post-processes one raw page file: re-encodes it to the requested format if the driver/decode
# handed back something else, crops it when asked, and appends the PAGE entry to the scan state.
# Shared by the WIA and the eSCL backends so both report the same fields.
function Finish-PageFile {
  param(
    [Parameter(Mandatory = $true)][string]$Part,
    [Parameter(Mandatory = $true)][hashtable]$State,
    [Parameter(Mandatory = $true)][string]$Want,
    [int]$Dpi = 300,
    [string]$Crop = 'off',
    [Parameter(Mandatory = $true)][string]$TempBase
  )
  $State.page++
  $State.sub = 'save'
  $target = Page-Path $TempBase $State.page
  $actual = Get-FileFormat $Part
  if (-not $State.nativeFormat -and $actual) { $State.nativeFormat = $actual }
  $converted = 0
  if ($actual -and $actual -ne $Want) {
    try {
      $saveDpi = $State.effectiveDpi
      if ($saveDpi -le 0) { $saveDpi = $Dpi }
      Convert-ImageFile $Part $Want $saveDpi | Out-Null
      $actual = $Want
      $converted = 1
    } catch {
      # Conversion refused: keep the device's bytes but stop calling the file by the wrong name.
      $target = [System.IO.Path]::ChangeExtension($target, '.' + $actual)
      Emit @('WARN', 'format', ('could not re-encode page ' + [string]$State.page + ' from ' + $actual + ', kept the device format'))
    }
  }
  if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
  Move-Item -LiteralPath $Part -Destination $target -Force

  if ($Crop -eq 'a4') {
    $cropDpi = $State.effectiveDpi
    if ($cropDpi -le 0) { $cropDpi = $Dpi }
    try {
      $cropRes = Crop-Image $target $cropDpi $Crop $Want
      if ($cropRes.dpiMismatch -and -not $State.cropDpiWarned) {
        $State.cropDpiWarned = $true
        Emit @('WARN', 'crop', ('the image header reports a different resolution than the negotiated ' + [string]$cropDpi + ' dpi; cropping with the negotiated value'))
      }
    } catch { Emit @('WARN', 'crop', ('could not crop page ' + [string]$State.page)) }
  }

  $bytes = 0
  if (Test-Path -LiteralPath $target) { $bytes = (Get-Item -LiteralPath $target).Length }
  $size = Get-ImageSize $target
  $dpiOut = $State.effectiveDpi
  if ($dpiOut -le 0) { $dpiOut = $size.dpi }
  [void]$State.pages.Add(@{
    temp = $target
    page = $State.page
    bytes = $bytes
    format = $actual
    width = $size.width
    height = $size.height
    dpi = $dpiOut
    converted = $converted
  })
}

# Opens one WIA session for exactly one sheet: connect, select the source, ask for the resolution,
# format and scan area, and hand the item back. A session is good for a single transfer only (this
# HP WSD driver answers a second transfer in the same session with "Unspecified error"), so the scan
# loop below opens a new session per page. Returns @{ dev; item; ok; error }.
function Open-ScanItem {
  param($chosen, [string]$src, [int]$dpi, [string]$formatId, [int]$intent, [string]$crop, $state)
  $session = @{ dev = $null; item = $null; ok = $false; error = '' }
  try {
    $session.dev = $chosen.info.Connect()
    if (-not $session.dev) { $session.error = 'the scanner could not be opened'; return $session }
    Set-Source $session.dev $src
    $count = 0
    try { $count = [int]$session.dev.Items.Count } catch { $count = 0 }
    if ($count -le 0) { $session.error = 'the scanner has no item to scan'; return $session }
    $session.item = $session.dev.Items.Item(1)
    if ($dpi -gt 0) {
      $pick = Select-Resolution $session.item $dpi
      $state.acceptedDpi = @($pick.accepted)
      if ($pick.effective -gt 0) { $state.effectiveDpi = [int]$pick.effective }
    }
    # Every sheet is a fresh item that starts at the driver defaults, so the format, the intent and the
    # scan area have to be asked for again on each one. This is what made the first feeder page work
    # and the second one fail with a bare "Unspecified error".
    Set-PropertyById $session.item.Properties 4106 $formatId | Out-Null
    Set-PropertyById $session.item.Properties 6146 $intent | Out-Null
    if ($crop -eq 'a4') {
      $extentDpi = $state.effectiveDpi
      if ($extentDpi -le 0) { $extentDpi = $dpi }
      Set-WiaA4Extent $session.item $extentDpi | Out-Null
    }
    $session.ok = $true
  } catch {
    $session.error = Get-ErrorText $_
    Release-ComObject $session.item
    Release-ComObject $session.dev
    $session.item = $null
    $session.dev = $null
  }
  return $session
}

function Close-ScanItem($session) {
  if (-not $session) { return }
  Release-ComObject $session.item
  Release-ComObject $session.dev
}

# Scans every page the current source hands over. Each page is written into the temporary directory
# first (see the note at the top); the returned state carries the page list so the caller can decide
# between a pdf and an image series. The state hashtable is passed in on purpose: it stays readable
# when a transfer fails, so the failure handler can report the number of completed pages.
function Invoke-ScanPages {
  param(
    $chosen,
    [string]$src,
    [string]$want,
    [string]$formatId,
    [int]$intent,
    [int]$dpi,
    [string]$work,
    [string]$crop,
    [int]$maxPages,
    $state
  )
  $session = $null
  $image = $null
  $tempBase = Join-Path $work ('page.' + $want)
  try {
    while ($state.page -lt $maxPages) {
      $image = $null
      $err = $null
      # Two attempts per sheet: the WSD driver sometimes refuses the very first transfer after it has
      # been idle, and a second session a moment later usually goes through.
      for ($attempt = 1; $attempt -le 2; $attempt++) {
        if ($attempt -gt 1) { Start-Sleep -Milliseconds 900 }
        Close-ScanItem $session
        $session = Open-ScanItem $chosen $src $dpi $formatId $intent $crop $state
        if (-not $session.ok) {
          $err = $session.error
          Close-ScanItem $session
          $session = $null
          continue
        }
        if (@($state.supported).Count -eq 0) {
          $state.supported = @(Get-SupportedFormats $session.item.Properties)
        }
        $state.sub = 'transfer'
        $transferErr = $null
        try { $image = $session.item.Transfer($formatId) } catch { $image = $null; $transferErr = $_ }
        if (-not $image) {
          # The driver ignored the requested format: take whatever it offers and re-encode afterwards.
          $secondErr = $null
          try { $image = $session.item.Transfer($FORMAT_IDS['bmp']) } catch { $image = $null; $secondErr = $_ }
          if ($secondErr) { $transferErr = $secondErr }
        }
        if (-not $image) {
          $thirdErr = $null
          try { $image = $session.item.Transfer() } catch { $image = $null; $thirdErr = $_ }
          if ($thirdErr) { $transferErr = $thirdErr }
        }
        Close-ScanItem $session
        $session = $null
        if ($image) { break }
        $err = $transferErr
        # An empty tray is the normal end of a feeder run, not a failure.
        if (Test-PaperEmpty $err) { $state.stoppedBy = 'paperEmpty'; break }
      }
      if ($state.stoppedBy -eq 'paperEmpty') { break }
      if (-not $image) {
        $state.stoppedBy = 'error'
        throw [System.InvalidOperationException]::new('the scanner refused to deliver an image: ' + (Get-ErrorText $err))
      }

      $part = (Page-Path $tempBase ($state.page + 1)) + '.part'
      if (Test-Path -LiteralPath $part) { Remove-Item -LiteralPath $part -Force }
      $image.SaveFile($part)
      Release-ComObject $image
      $image = $null
      Finish-PageFile -Part $part -State $state -Want $want -Dpi $dpi -Crop $crop -TempBase $tempBase

      # The flatbed holds one sheet: one page is the whole scan.
      if ($src -ne 'feed') { break }
    }
    if ($state.page -ge $maxPages) { $state.stoppedBy = 'maxPages' }
  } finally {
    Release-ComObject $image
    Close-ScanItem $session
  }
}

# Scans over eSCL. The ADF stack is ONE job: every NextDocument call hands back the next sheet and a
# 404 means the stack is finished. That is why a multi-page feeder scan works here while the WIA path
# of this driver stops after one sheet.
function Invoke-EsclScanPages {
  param(
    [Parameter(Mandatory = $true)][string]$Base,
    $Caps,
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Want,
    [string]$Mode = 'color',
    [int]$Dpi = 300,
    [string]$Crop = 'off',
    [int]$MaxPages = 50,
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][hashtable]$State
  )
  $tempBase = Join-Path $Work ('page.' + $Want)
  # The InputSource token this firmware wants for the document feeder is 'Feeder'; it answers
  # "409 Conflict" to the spec-style 'ADF' token, which looks like a device error but is only a
  # naming mismatch.
  $input = 'Platen'
  if ($Src -eq 'feed') { $input = 'Feeder' }

  $status = Get-EsclStatus $Base
  if ($input -eq 'Feeder' -and $status.ok -and $status.adfState -eq 'ScannerAdfEmpty') {
    # Same signal as WIA_ERROR_PAPER_EMPTY: the feeder is empty, and 'auto' turns it into a flatbed
    # run a moment later.
    $State.stoppedBy = 'paperEmpty'
    return
  }

  $accepted = @()
  if ($input -eq 'ADF') { $accepted = @($Caps.adfResolutions) } else { $accepted = @($Caps.platenResolutions) }
  $State.acceptedDpi = $accepted
  $effective = Select-EsclResolution -Accepted $accepted -Want $Dpi
  if ($effective -le 0) { $effective = $Dpi }
  $State.effectiveDpi = [int]$effective
  $State.supported = @('jpeg')
  $State.nativeFormat = 'jpeg'

  $colorMode = 'RGB24'
  if ($Mode -eq 'gray' -or $Mode -eq 'grey') { $colorMode = 'Grayscale8' }
  elseif ($Mode -eq 'text') {
    $colorMode = 'Grayscale8'
    Emit @('WARN', 'mode', 'this device has no eSCL text mode; the page was scanned in grayscale')
  }

  $xml = New-EsclSettingsXml -InputSource $input -Resolution $effective -ColorMode $colorMode -Intent 'Document'
  $job = New-EsclJob -Base $Base -SettingsXml $xml
  if (-not $job.ok) {
    # The feeder refuses the job with 409 when it has nothing to feed, so an empty tray ends the run
    # instead of being reported as a broken scanner.
    if ($input -eq 'Feeder' -and $job.status -eq 409) {
      $State.stoppedBy = 'paperEmpty'
      return
    }
    $State.stoppedBy = 'error'
    throw [System.InvalidOperationException]::new($job.error)
  }
  $retries = 0
  try {
    while ($State.page -lt $MaxPages) {
      $raw = Join-Path $Work ('escl-' + [string]($State.page + 1) + '.jpg')
      $doc = Get-EsclDocument -JobUrl $job.jobUrl -OutFile $raw
      if ($doc.ok) {
        Finish-PageFile -Part $raw -State $State -Want $Want -Dpi $Dpi -Crop $Crop -TempBase $tempBase
        # The flatbed holds a single sheet; a feeder keeps going until the job says it is empty.
        if ($input -ne 'Feeder') { break }
        continue
      }
      if ($doc.gone) {
        # 404/410 from NextDocument is the normal end of a feeder batch, not an error.
        if ($input -eq 'Feeder') { $State.stoppedBy = 'paperEmpty' }
        break
      }
      # 503 while the scanner warms up for the first sheet: the job is fine, just not ready yet.
      if (($doc.status -eq 409 -or $doc.status -eq 503) -and $retries -lt 60) {
        $retries++
        Start-Sleep -Milliseconds 1000
        continue
      }
      $State.stoppedBy = 'error'
      throw [System.InvalidOperationException]::new('the scanner refused to deliver an image over eSCL: ' + $doc.error + ' (HTTP ' + [string]$doc.status + ')')
    }
    if ($State.page -ge $MaxPages) { $State.stoppedBy = 'maxPages' }
    # A page batch that was not read to its natural end leaves a job in the device queue; canceled
    # jobs pile up and the scanner starts answering 409 to new jobs. Reading once more lets the
    # device finish the job on its own (404), which keeps the queue clean.
    $drain = Get-EsclDocument -JobUrl $job.jobUrl -OutFile (Join-Path $Work 'drain.bin') -TimeoutSec 60
    if (-not $drain.ok -and -not $drain.gone) { Start-Sleep -Milliseconds 500 }
  } finally {
    Remove-EsclJob $job.jobUrl
  }
}

# Everything that can be validated without touching a device is validated first.
$fmt = 'jpg'
$pdfMode = $false
if ($Format) {
  $fmt = $Format.ToLowerInvariant()
  if ($fmt -eq 'pdf') { $pdfMode = $true; $fmt = 'jpg' }
}
if (-not $FORMAT_IDS.ContainsKey($fmt)) { Fail ('unknown format: ' + $Format) 'unknownFormat' }
$want = $fmt
if ($want -eq 'jpeg') { $want = 'jpg' }
if ($want -eq 'tiff') { $want = 'tif' }
$formatId = $FORMAT_IDS[$fmt]

$srcRequest = 'auto'
if ($Source) { $srcRequest = $Source.ToLowerInvariant() }
if (@('auto', 'flat', 'feed') -notcontains $srcRequest) { Fail ('unknown source: ' + $Source) 'unknownSource' }

$cropMode = 'off'
if ($Crop) { $cropMode = $Crop.ToLowerInvariant() }
if (@('off', 'a4') -notcontains $cropMode) { Fail ('unknown crop: ' + $Crop) 'unknownCrop' }

$intent = 1
switch ($Mode) {
  'color' { $intent = 1 }
  'gray'  { $intent = 2 }
  'grey'  { $intent = 2 }
  'text'  { $intent = 4 }
  default { Fail ('unknown mode: ' + $Mode) 'unknownMode' }
}

$base = [System.IO.Path]::GetFullPath($Output)
if ($pdfMode) { $base = [System.IO.Path]::ChangeExtension($base, '.pdf') }
$baseDir = [System.IO.Path]::GetDirectoryName($base)
if ($baseDir -and -not (Test-Path -LiteralPath $baseDir)) {
  New-Item -ItemType Directory -Path $baseDir -Force | Out-Null
}
$imgBase = $base
if ($pdfMode) { $imgBase = [System.IO.Path]::ChangeExtension($base, '.' + $want) }

$work = Join-Path $env:TEMP ('dsh-scan-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -Force | Out-Null

$wia = $null
$chosen = $null
$probe = $null
$state = $null
$phase = 'open'
try {
  $wia = Get-WiaDevices
  if (-not $wia.ok) { Fail ('cannot open the WIA device manager: ' + $wia.error) 'scannerUnreadable' }
  if (@($wia.devices).Count -eq 0) { Fail 'no WIA scanner found on this computer' 'noScanner' }

  if (-not $Device) {
    # No device named: pick the first device that can actually be opened. Some all-in-one drivers
    # register a second, permanently off-line eSCL entry that sorts before the real WIA device;
    # taking the first entry blindly turned every default scan into a device error.
    $failures = @()
    foreach ($d in $wia.devices) {
      $open = Open-WiaDevice $d
      if ($open.ok) { $chosen = $d; $probe = $open.device; break }
      $failures += ([string]$d.name + ': ' + [string]$open.error)
    }
    if (-not $chosen) { Fail ('no scanner could be opened: ' + ($failures -join '; ')) 'deviceError' }
  } else {
    # Exact name / DeviceID / index first, then a unique name prefix (the printer tools accept a
    # prefix too, and typing the whole "(HP Color LaserJet MFP M283cdw)" is a natural mistake).
    $found = Find-WiaDevice $wia.devices $Device
    if (-not $found.device -and @($found.ambiguous).Count -gt 0) {
      Fail ('more than one scanner matches "' + $Device + '": ' + ([string]::Join(', ', @($found.names))) + '. Give the full name.') 'scannerAmbiguous'
    }
    $chosen = $found.device
    if (-not $chosen) { Fail (Get-WiaNotFoundText $Device $wia.devices) 'scannerNotFound' }
  }
  $chosenName = [string]$chosen.name

  # Document handling: 3086 capabilities (bit 1 = feeder, bit 2 = flatbed, bit 4 = duplex),
  # 3088 selection, 3087 status is deliberately not used (several drivers always report 1).
  $phase = 'device'
  $caps = 0
  if (-not $probe) { $probe = $chosen.info.Connect() }
  $capsRaw = Get-PropertyById $probe.Properties 3086
  if ($null -ne $capsRaw) { try { $caps = [int]$capsRaw } catch { $caps = 0 } }
  $hasFeeder = (($caps -band 1) -eq 1)
  $hasFlatbed = (($caps -band 2) -eq 2)
  if ($srcRequest -eq 'feed' -and -not $hasFeeder) { Fail 'this scanner has no document feeder' 'noFeeder' }
  if ($srcRequest -eq 'flat' -and -not $hasFlatbed) { Fail 'this scanner has no flatbed' 'noFlatbed' }

  # auto must not trust 3088: it holds whatever the driver used last, so a loaded feeder would
  # still be scanned on the glass and come back blank. Prefer the feeder, fall back to the glass.
  $src = $srcRequest
  if ($srcRequest -eq 'auto') {
    if ($hasFeeder) { $src = 'feed' } else { $src = 'flat' }
  }

  # Some drivers default DPS_PAGES to 1, which stops the ADF after the first sheet; ask for all
  # The device has to be open before the driver publishes its network address (WIA property 38919,
  # "Presentation URL"), so read that address now and only then let go of the probe connection.
  $esclHint = ''
  foreach ($id in 38919, 38920) {
    $value = $null
    try { $value = Get-PropertyById $probe.Properties $id } catch { $value = $null }
    if ($value -and ([string]$value -match '^(https?)://')) { $esclHint = [string]$value; break }
  }
  # The capability probe has done its job; the scan backends open their own connections, so let go
  # of this one instead of holding two at the same time.
  Release-ComObject $probe
  $probe = $null

  # Prefer eSCL when the device answers it: a network scanner runs a whole ADF stack as one job,
  # while the WIA path of this driver stops after the first sheet. The WIA path stays for USB-only
  # hardware and for devices that do not speak eSCL.
  $esclBase = ''
  $esclCaps = $null
  $esclProbe = Resolve-EsclBase $esclHint
  if (-not $esclProbe.ok) { $esclProbe = Get-EsclBaseFromWiaDevice $chosen.info }
  if ($esclProbe.ok) {
    $esclBase = [string]$esclProbe.base
    $esclCaps = Get-EsclCapabilities $esclBase
    Emit @('WARN', 'backend', ('using the eSCL network interface at ' + $esclBase))
  }

  $phase = 'scan'
  $state = @{
    page = 0
    stoppedBy = 'complete'
    effectiveDpi = 0
    acceptedDpi = @()
    supported = @()
    nativeFormat = ''
    cropDpiWarned = $false
    pages = New-Object System.Collections.ArrayList
  }
  if ($esclBase) {
    Invoke-EsclScanPages -Base $esclBase -Caps $esclCaps -Src $src -Want $want -Mode $Mode -Dpi $Dpi -Crop $cropMode -MaxPages $MAX_PAGES -Work $work -State $state
  } else {
    Invoke-ScanPages -chosen $chosen -src $src -want $want -formatId $formatId -intent $intent -dpi $Dpi -work $work -crop $cropMode -maxPages $MAX_PAGES -state $state
  }
  $fallback = ''
  if ($state.page -eq 0 -and $state.stoppedBy -eq 'paperEmpty' -and $srcRequest -eq 'auto' -and $hasFlatbed) {
    Emit @('WARN', 'feederEmpty', 'the document feeder is empty, falling back to the flatbed')
    $src = 'flat'
    $state = @{
      page = 0
      stoppedBy = 'complete'
      effectiveDpi = 0
      acceptedDpi = @()
      supported = @()
      nativeFormat = ''
      cropDpiWarned = $false
      pages = New-Object System.Collections.ArrayList
    }
    if ($esclBase) {
      Invoke-EsclScanPages -Base $esclBase -Caps $esclCaps -Src $src -Want $want -Mode $Mode -Dpi $Dpi -Crop $cropMode -MaxPages $MAX_PAGES -Work $work -State $state
    } else {
      Invoke-ScanPages -chosen $chosen -src $src -want $want -formatId $formatId -intent $intent -dpi $Dpi -work $work -crop $cropMode -maxPages $MAX_PAGES -state $state
    }
    $fallback = 'feederEmpty'
  }

  if ($state.page -eq 0) {
    if ($state.stoppedBy -eq 'paperEmpty') {
      if ($hasFlatbed) { Fail 'the document feeder is empty (load the feeder, or ask for the flatbed)' 'feederEmpty' }
      Fail 'the document feeder is empty' 'feederEmpty'
    }
    Fail 'the scanner produced no page' 'noPage'
  }

  $pdfPath = ''
  if ($pdfMode) {
    $phase = 'pdf'
    $pdfDpi = $state.effectiveDpi
    if ($pdfDpi -le 0) { $pdfDpi = $Dpi }
    $images = @()
    foreach ($p in $state.pages) { $images += $p.temp }
    $pdfBytes = Write-Pdf -images $images -output $base -dpi $pdfDpi
    $pdfPath = $base
    Emit @('PDF', $base, [string]$pdfBytes, [string]$state.page, [string]$pdfDpi)
  }

  $phase = 'save'
  Move-ScanPages -state $state -imgBase $imgBase -pdfMode $pdfMode -keepPages ([bool]$KeepPages) -partial $false

  $sourceText = 'flatbed'
  if ($src -eq 'feed') { $sourceText = 'feeder' }
  Emit @('OK', 'scan', $chosenName, [string]$state.page, [string]$Dpi, [string]$state.effectiveDpi, $sourceText, [string]::Join(',', $state.supported), [string]::Join(',', $state.acceptedDpi), $state.stoppedBy, [string]$caps, $state.nativeFormat, $cropMode, $fallback)
} catch {
  # Completed pages are not thrown away: they are moved out as an image series, with PAGE records,
  # and the STAT line tells the host how far the scan got.
  if ($state -and $state.page -gt 0) {
    $st = $state
    try {
      Emit @('WARN', 'partial', ('the scan stopped after ' + [string]$st.page + ' page(s); they were kept as separate images'))
      Move-ScanPages -state $st -imgBase $imgBase -pdfMode $pdfMode -keepPages $true -partial $true
    } catch { }
    # stoppedBy must not claim 'complete' on a failed run; the device and source are repeated here
    # because the OK record is only emitted on a clean finish.
    $stopReason = $st.stoppedBy
    if (-not $stopReason -or $stopReason -eq 'complete') { $stopReason = 'error' }
    $devName = ''
    if ($chosenName) { $devName = [string]$chosenName }
    $srcText = ''
    if ($src -eq 'feed') { $srcText = 'feeder' } elseif ($src -eq 'flat') { $srcText = 'flatbed' }
    try { Emit @('STAT', [string]$st.page, $stopReason, $devName, $srcText) } catch { }
  }
  $code = 'deviceError'
  switch ($phase) {
    'scan' {
      # A page that could not be written to the temporary directory is a save failure, not a
      # transfer failure: the scanner delivered, the disk refused.
      if ($state -and $state.sub -eq 'save') { $code = 'saveFailed' } else { $code = 'transferFailed' }
    }
    'pdf'  {
      # A4 300 dpi 彩色页约 1-3 MB; a long multi-page scan can hit the PDF size cap, which is a
      # different remedy (split the scan / lower the resolution) from a broken PDF writer.
      $msg = ''
      if ($_ -and $_.Exception -and $_.Exception.Message) { $msg = [string]$_.Exception.Message }
      if ($msg -match 'byte limit') { $code = 'pdfTooLarge' } else { $code = 'pdfFailed' }
    }
    'save' { $code = 'saveFailed' }
    'device' { $code = 'deviceError' }
    default { $code = 'deviceError' }
  }
  Emit @('CODE', $code)
  Emit @('ERR', $_.Exception.Message)
  exit 1
} finally {
  Release-ComObject $probe
  if ($wia) {
    if ($wia.infos) { Release-ComObject $wia.infos }
    if ($wia.manager) { Release-ComObject $wia.manager }
  }
  if ($work -and (Test-Path -LiteralPath $work)) {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
  }
}

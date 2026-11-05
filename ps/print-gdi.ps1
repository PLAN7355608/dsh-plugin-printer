# print-gdi.ps1 - GDI+ print backend, dot-sourced by print-file.ps1 (not run directly).
#
# Why this backend exists: the shell "print" verb hands the whole job to the associated program, so
# the plugin has no say in the page range, the copy order or duplexing, and the spooler text path
# cannot do it either. System.Drawing's PrintDocument renders the pages inside this process instead,
# which lets one spool job carry an exact page selection, an exact copy order (collated or stacked)
# and a per-job duplex request - without touching the printer's default preferences.
#
# Only text and image files are rendered here. PDF/OOXML/Office files need their own renderer (WPS,
# Office, a PDF reader) and stay on the verb path, which cannot honour these options.
#
# Pure ASCII on purpose: Windows PowerShell 5.1 reads .ps1 files as ANSI.

$script:GdiReady = $false
try {
  Add-Type -AssemblyName System.Drawing -ErrorAction Stop
  $null = [System.Drawing.Printing.PrintDocument]
  $script:GdiReady = $true
} catch {
  $script:GdiReady = $false
}

# A caught error must never come out as an empty string.
function Gdi-ErrorText($err) {
  $parts = @()
  if ($err.Exception -and $err.Exception.Message) { $parts += [string]$err.Exception.Message }
  if ($err.Exception.InnerException -and $err.Exception.InnerException.Message) { $parts += [string]$err.Exception.InnerException.Message }
  if (-not $parts.Count -and $err.CategoryInfo -and $err.CategoryInfo.Reason) { $parts += [string]$err.CategoryInfo.Reason }
  if (-not $parts.Count -and $err.FullyQualifiedErrorId) { $parts += [string]$err.FullyQualifiedErrorId }
  if (-not $parts.Count) { $parts += 'unknown failure' }
  return ($parts -join ' / ')
}

function Test-IsImage([string]$path) {
  $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
  return (@('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff') -contains $ext)
}

function Test-BytePrefix([byte[]]$bytes, [byte[]]$prefix) {
  if ($bytes.Length -lt $prefix.Length) { return $false }
  for ($i = 0; $i -lt $prefix.Length; $i++) { if ($bytes[$i] -ne $prefix[$i]) { return $false } }
  return $true
}

# Everything that is not an image is treated as text, so the content is checked before it is
# decoded: a PDF, an Office document or an archive read as text turns into thousands of pages of
# binary garbage in the spooler (a 15-page report once reached the queue as a 932-page job).
# A NUL byte in a file without a text BOM, or a known binary signature, is enough to refuse it;
# a BOM is checked first so UTF-16 files are not mistaken for binary.
function Read-HeadBytes([string]$path, [int]$max) {
  $probe = @{ ok = $false; bytes = @(); error = '' }
  $stream = $null
  try {
    $stream = [System.IO.File]::OpenRead($path)
    $len = [int][Math]::Min($max, $stream.Length)
    if ($len -lt 1) { $probe.ok = $true; return $probe }
    $buffer = New-Object byte[] $len
    $read = 0
    # One Read() may return fewer bytes than asked for, and the old code dropped the count too:
    # keep reading until the buffer is full or the file ends.
    while ($read -lt $len) {
      $n = $stream.Read($buffer, $read, $len - $read)
      if ($n -le 0) { break }
      $read += $n
    }
    if ($read -lt $len) {
      $trimmed = New-Object byte[] $read
      [Array]::Copy($buffer, $trimmed, $read)
      $buffer = $trimmed
    }
    $probe.ok = $true
    $probe.bytes = $buffer
    return $probe
  } catch {
    $probe.error = Gdi-ErrorText $_
    return $probe
  } finally {
    if ($stream) { $stream.Dispose() }
  }
}

# A byte-order mark is proof that the file is text, and it is checked BEFORE the NUL scan: UTF-16
# text is full of NUL bytes, so the old order refused every UTF-16 file and made the documented
# encoding=unicode option unreachable.
function Get-TextBom($bytes) {
  $b = @($bytes)
  if ($b.Count -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { return 'utf8bom' }
  if ($b.Count -ge 4 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE -and $b[2] -eq 0x00 -and $b[3] -eq 0x00) { return 'utf32le' }
  if ($b.Count -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) { return 'utf16le' }
  if ($b.Count -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF) { return 'utf16be' }
  return ''
}

# Returns @{ kind = 'text' | 'binary' | 'unreadable'; bytes; bom; error }. Three outcomes, because
# "the file could not be opened" must never be read as "text": a directory or a locked file used to
# slip through and be decoded. A binary signature, or a NUL byte in a file without a text BOM, is
# enough to refuse it.
function Get-TextProbe([string]$path) {
  $probe = @{ kind = 'text'; bytes = @(); bom = ''; error = '' }
  $head = Read-HeadBytes $path 8192
  if (-not $head.ok) {
    $probe.kind = 'unreadable'
    $probe.error = $head.error
    return $probe
  }
  $bytes = $head.bytes
  $probe.bytes = $bytes
  $probe.bom = Get-TextBom $bytes
  if ($probe.bom) { return $probe }
  if (@($bytes).Count -lt 1) { return $probe }
  foreach ($b in @($bytes)) { if ($b -eq 0) { $probe.kind = 'binary'; return $probe } }
    $magics = @(
      @(0x25, 0x50, 0x44, 0x46, 0x2D),      # %PDF-
      @(0x50, 0x4B),                        # PK   (zip, OOXML, jar)
      @(0xD0, 0xCF, 0x11, 0xE0),            # OLE  (doc/xls/ppt)
      @(0x7B, 0x5C, 0x72, 0x74, 0x66),      # {\rtf
      @(0x1F, 0x8B),                        # gzip
      @(0x52, 0x61, 0x72, 0x21),            # Rar!
      @(0x37, 0x7A, 0xBC, 0xAF),            # 7z
      @(0x50, 0x4B, 0x03, 0x04),            # zip local header
      @(0x4D, 0x5A),                        # MZ   (exe/dll)
      @(0x7F, 0x45, 0x4C, 0x46)             # ELF
    )
  foreach ($m in $magics) { if (Test-BytePrefix $bytes ([byte[]]$m)) { $probe.kind = 'binary'; return $probe } }
  return $probe
}

# Kept for the regression guard, which asserts on the yes/no answer only. An unreadable file counts
# as "not text" here so a caller that only checks this flag still refuses it.
function Test-LooksBinary([string]$path) {
  return ((Get-TextProbe $path).kind -ne 'text')
}

# Picks a font that can draw CJK text: if the requested family is not installed GDI+ silently
# substitutes another one, so the resolved name is checked before accepting it.
# The spooler display name and the GDI queue name can differ ("Microsoft XPS Document Writer v4"
# in the registry vs "Microsoft XPS Document Writer" for PrinterSettings), so names are compared
# with a trailing " vN" stripped before giving up.
# Reads a text file with the same encoding rules as the queue path. It lives here (not in
# print-file.ps1) because Invoke-GdiPrint is the only caller inside the renderer, and the GDI
# backend must work when it is dot-sourced on its own - that is what the regression test does.
function Read-TextSmart([string]$path, [string]$encoding) {
  $bytes = [System.IO.File]::ReadAllBytes($path)
  if ($encoding -eq 'utf8') { return ([System.Text.Encoding]::UTF8.GetString($bytes)) }
  if ($encoding -eq 'gbk') { return ([System.Text.Encoding]::GetEncoding(936).GetString($bytes)) }
  if ($encoding -eq 'unicode') { return ([System.Text.Encoding]::Unicode.GetString($bytes)) }
  if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    return ([System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3))
  }
  if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
    return ([System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2))
  }
  if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
    return ([System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2))
  }
  $strict = New-Object System.Text.UTF8Encoding $false, $true
  try { return ($strict.GetString($bytes)) } catch { return ([System.Text.Encoding]::GetEncoding(936).GetString($bytes)) }
}

function Normalize-PrinterName([string]$name) {
  $n = ''
  if ($name) { $n = $name.Trim() }
  return ($n -replace '(?i)\s+v[0-9]+$', '').ToLowerInvariant()
}

function Resolve-GdiPrinterName([string]$printer, $installed) {
  foreach ($p in $installed) { if ([string]$p -eq $printer) { return [string]$p } }
  $want = Normalize-PrinterName $printer
  foreach ($p in $installed) { if ((Normalize-PrinterName ([string]$p)) -eq $want) { return [string]$p } }
  foreach ($p in $installed) { if ((Normalize-PrinterName ([string]$p)).Contains($want)) { return [string]$p } }
  return ''
}

function Get-GdiFont([single]$size) {
  foreach ($name in @('Microsoft YaHei', 'SimSun', 'NSimSun', 'Segoe UI', 'Arial')) {
    try {
      $f = New-Object System.Drawing.Font($name, $size)
      if ($f.Name -eq $name) { return $f }
      $f.Dispose()
    } catch { }
  }
  return (New-Object System.Drawing.Font([System.Drawing.FontFamily]::GenericMonospace, $size))
}

# Parses a page selection: "", all, *, odd, even, 3, 1-3, 5-, -4, and comma lists of those.
# Returns @{ pages = int[]; label = '...'; error = '...' }.
function Resolve-PageRange([string]$spec, [int]$total) {
  $info = @{ pages = @(); label = ''; error = '' }
  $all = New-Object System.Collections.ArrayList
  for ($i = 1; $i -le $total; $i++) { [void]$all.Add($i) }
  $s = ''
  if ($spec) { $s = $spec.Trim().ToLowerInvariant() }
  if (-not $s -or $s -eq 'all' -or $s -eq '*') { $info.pages = @($all); $info.label = 'all'; return $info }
  if ($s -eq 'odd' -or $s -eq 'even') {
    $picked = New-Object System.Collections.ArrayList
    foreach ($n in $all) {
      $isOdd = (($n % 2) -eq 1)
      if (($s -eq 'odd') -eq $isOdd) { [void]$picked.Add($n) }
    }
    if ($picked.Count -eq 0) { $info.error = 'the page range "' + $s + '" selected no pages'; return $info }
    $info.pages = @($picked)
    $info.label = $s
    return $info
  }
  $seen = @{}
  $picked = New-Object System.Collections.ArrayList
  foreach ($part in @($s -split ',')) {
    $t = $part.Trim()
    if (-not $t) { continue }
    $from = 0
    $to = 0
    if ($t -match '^([0-9]+)$') {
      $from = [int]$Matches[1]
      $to = $from
    } elseif ($t -match '^([0-9]+)\s*-\s*([0-9]*)$') {
      $from = [int]$Matches[1]
      if ($Matches[2] -eq '') { $to = $total } else { $to = [int]$Matches[2] }
    } elseif ($t -match '^-\s*([0-9]+)$') {
      $from = 1
      $to = [int]$Matches[1]
    } else {
      $info.error = 'cannot read the page range "' + $t + '": use forms like 1-3,5,8- or odd/even'
      return $info
    }
    if ($from -lt 1) { $info.error = 'page numbers start at 1 (got "' + $t + '")'; return $info }
    if ($to -lt $from) { $info.error = 'the page range "' + $t + '" ends before it starts'; return $info }
    if ($to -gt $total) { $info.error = 'page ' + $to + ' is out of range: this document has ' + $total + ' page(s)'; return $info }
    for ($i = $from; $i -le $to; $i++) {
      if (-not $seen.ContainsKey($i)) { $seen[$i] = $true; [void]$picked.Add($i) }
    }
  }
  if ($picked.Count -eq 0) { $info.error = 'the page range selected no pages'; return $info }
  $info.pages = @($picked)
  $info.label = (@($picked) -join ',')
  return $info
}

# Wraps one source line into as many display lines as the printable width needs. Without this the
# renderer drew long lines straight off the right edge, and the page count was computed from the
# unwrapped text, so the overflow was simply lost.
function Get-WrappedLines([string]$line, $font, $graphics, [single]$maxWidth) {
  $out = New-Object System.Collections.ArrayList
  $rest = [string]$line
  if (-not $graphics -or -not $font -or $maxWidth -le 0) { [void]$out.Add($rest); return ,$out }
  while ($true) {
    if ($rest.Length -lt 1) { [void]$out.Add(''); break }
    $w = [single]$graphics.MeasureString($rest, $font).Width
    if ($w -le $maxWidth) { [void]$out.Add($rest); break }
    # Longest prefix that still fits, found by binary search over the character count.
    $lo = 1
    $hi = $rest.Length - 1
    $fit = 0
    while ($lo -le $hi) {
      $mid = [int][Math]::Floor(($lo + $hi) / 2)
      $mw = [single]$graphics.MeasureString($rest.Substring(0, $mid), $font).Width
      if ($mw -le $maxWidth) { $fit = $mid; $lo = $mid + 1 } else { $hi = $mid - 1 }
    }
    if ($fit -lt 1) { $fit = 1 }
    [void]$out.Add($rest.Substring(0, $fit))
    $rest = $rest.Substring($fit)
  }
  return ,$out
}

# Splits text into pages of display lines. A form feed always starts a new page; the rest is paginated
# by the number of text lines that fit between the printer's margins.
function Get-TextPages([string]$text, [single]$lineHeight, [single]$usableHeight, $font, $graphics, [single]$usableWidth) {
  $perPage = 1
  if ($lineHeight -gt 0 -and $usableHeight -gt 0) { $perPage = [int][Math]::Floor($usableHeight / $lineHeight) }
  if ($perPage -lt 1) { $perPage = 1 }
  $pages = New-Object System.Collections.ArrayList
  $segments = @($text -split ([string][char]12))
  for ($s = 0; $s -lt $segments.Count; $s++) {
    $segment = [string]$segments[$s]
    # An empty trailing segment comes from a final form feed and must not become a blank page.
    if ($segment.Length -eq 0 -and $segments.Count -gt 1) { continue }
    $lines = New-Object System.Collections.ArrayList
    foreach ($source in @($segment -split "\r\n|\n|\r")) {
      foreach ($display in @(Get-WrappedLines ([string]$source) $font $graphics $usableWidth)) {
        [void]$lines.Add([string]$display)
        if ($lines.Count -ge $perPage) {
          [void]$pages.Add(@($lines))
          $lines = New-Object System.Collections.ArrayList
        }
      }
    }
    # Keep a partial page, but never append an empty one after an exact multiple: two lines with
    # room for two used to produce a second, blank page.
    if ($lines.Count -gt 0) { [void]$pages.Add(@($lines)) }
  }
  if ($pages.Count -eq 0) { [void]$pages.Add(@()) }
  return ,$pages
}

function Get-ImageFrameCount([string]$path) {
  $img = [System.Drawing.Image]::FromFile($path)
  try { return [int]$img.GetFrameCount([System.Drawing.Imaging.FrameDimension]::Page) }
  finally { $img.Dispose() }
}

# Shortens a long page list for the record line (99 copies of a long document would blow it up).
function Format-PageList($list) {
  $items = @($list)
  if ($items.Count -le 40) { return ($items -join ',') }
  return (($items[0..19] -join ',') + ',...(' + $items.Count + ' total)...,' + ($items[($items.Count - 10)..($items.Count - 1)] -join ','))
}

# Renders one spool job with GDI+. Returns a hashtable; the caller writes the record line.
function Invoke-GdiPrint {
  param(
    [string]$Path,
    [string]$Printer,
    [string]$PageRange = 'all',
    [int]$Copies = 1,
    [string]$Collate = 'true',
    [string]$Duplex = 'off',
    [string]$OutputFile = '',
    [string]$TextEncoding = 'auto',
    [int]$MaxPages = 200,
    [switch]$AllowHugeJob,
    [switch]$PlanOnly
  )
  $result = @{
    ok = $false; error = ''; code = ''; kind = ''; pagesTotal = 0; pagesPrinted = ''
    copies = $Copies; collate = 'true'; duplex = $Duplex; order = ''; output = ''; plan = ''
    printer = $Printer; renamed = $false
  }
  if ($Copies -lt 1) { $Copies = 1 }
  $result.copies = $Copies
  if (-not $script:GdiReady) {
    $result.error = 'the GDI+ print backend is unavailable: System.Drawing could not be loaded'
    $result.code = 'gdiUnavailable'
    return $result
  }
  if (-not (Test-Path -LiteralPath $Path)) {
    $result.error = 'file not found: ' + $Path
    $result.code = 'fileNotFound'
    return $result
  }
  $installed = @()
  try { $installed = @([System.Drawing.Printing.PrinterSettings]::InstalledPrinters) } catch { $installed = @() }
  $gdiPrinter = $Printer
  if ($installed.Count -gt 0) {
    $gdiPrinter = Resolve-GdiPrinterName $Printer $installed
    if (-not $gdiPrinter) {
      $show = @($installed | Select-Object -First 10)
      $result.error = 'no printer named "' + $Printer + '" is installed. Available: ' + ($show -join ', ') + '.'
      $result.code = 'unknownPrinter'
      return $result
    }
    $result.printer = $gdiPrinter
    $result.renamed = ($gdiPrinter -ne $Printer)
  }
  # NOTE: the System.Drawing.Printing.Duplex names describe the FLIP AXIS, not the binding edge,
  # so they read backwards for paper the user drops in portrait: 'Vertical' is long-edge (book)
  # binding and 'Horizontal' is short-edge binding. Verified against this machine's HP V4 driver by
  # setting TwoSidedLongEdge / TwoSidedShortEdge in the printing preferences and reading the .NET
  # value back: TwoSidedLongEdge -> Duplex.Vertical (2), TwoSidedShortEdge -> Duplex.Horizontal (3).
  # The 1.2.0 build mapped long->Horizontal, which made the back of a long-edge flip upside down.
  $duplexMap = @{ 'off' = 'Simplex'; 'long' = 'Vertical'; 'short' = 'Horizontal' }
  if (-not $duplexMap.ContainsKey($Duplex)) {
    $result.error = 'unknown duplex mode: ' + $Duplex
    $result.code = 'unknownDuplex'
    return $result
  }
  $duplexName = [string]$duplexMap[$Duplex]
  $collateText = 'true'
  if ($Collate -and $Collate.Trim().ToLowerInvariant() -eq 'false') { $collateText = 'false' }
  $result.collate = $collateText
  $result.duplex = $Duplex

  $doc = $null
  $font = $null
  $measure = $null
  $image = $null
  try {
    $doc = New-Object System.Drawing.Printing.PrintDocument
    $doc.PrinterSettings.PrinterName = $gdiPrinter
    $doc.DocumentName = [System.IO.Path]::GetFileName($Path)
    $isImage = Test-IsImage $Path
    $pages = @()
    $total = 0
    $lineHeight = [single]0
    if ($isImage) {
      $total = Get-ImageFrameCount $Path
      $result.kind = 'image'
    } else {
      # The probe tells "text" from "binary" from "cannot be read"; the last one used to be treated
      # as text, so a directory or a locked file reached the text decoder.
      $probe = Get-TextProbe $Path
      if ($probe.kind -eq 'unreadable') {
        $result.error = 'the file could not be read: ' + $probe.error
        $result.code = 'fileUnreadable'
        return $result
      }
      if ($probe.kind -eq 'binary') {
        $result.error = 'this file is not text: it starts with a binary signature (PDF, Office document, archive or executable) or contains NUL bytes without a text byte-order mark, so the text renderer would decode it and print thousands of pages of garbage. Print it with the shell print verb instead, or convert it to text or an image first.'
        $result.code = 'notText'
        return $result
      }
      $font = Get-GdiFont 10
      $measure = $doc.PrinterSettings.CreateMeasurementGraphics()
      if (-not $measure) { throw 'the printer driver did not report a printable area' }
      $lineHeight = [single]$font.GetHeight($measure)
      $usable = [single]$doc.DefaultPageSettings.PrintableArea.Height
      $usableWidth = [single]$doc.DefaultPageSettings.PrintableArea.Width
      $text = Read-TextSmart $Path $TextEncoding
      $pages = Get-TextPages $text $lineHeight $usable $font $measure $usableWidth
      $total = $pages.Count
      $result.kind = 'text'
    }
    if ($total -lt 1) {
      $result.error = 'the file produced no printable pages'
      $result.code = 'noPrintablePages'
      return $result
    }
    $sel = Resolve-PageRange $PageRange $total
    if ($sel.error) {
      $result.error = $sel.error
      $result.code = 'badPageRange'
      return $result
    }
    # The fuse counts the pages this job would actually print, not the size of the file: a page
    # range of 1-2 out of a 5000-page log is a two-page job, and refusing it was a false alarm.
    $selectedCount = @($sel.pages).Count
    if (-not $AllowHugeJob -and $selectedCount -gt $MaxPages) {
      $result.error = ('this job would print ' + $selectedCount + ' pages (' + $total + ' in the file), far more than expected: the ' + $MaxPages + '-page safety limit keeps a mis-detected file from emptying a paper tray. Check the file, then pass allowHugeJob to print it anyway.')
      $result.code = 'tooManyPages'
      $result.pagesTotal = $total
      return $result
    }
    $order = New-Object System.Collections.ArrayList
    if ($collateText -eq 'false') {
      foreach ($p in $sel.pages) { for ($c = 0; $c -lt $Copies; $c++) { [void]$order.Add($p) } }
    } else {
      for ($c = 0; $c -lt $Copies; $c++) { foreach ($p in $sel.pages) { [void]$order.Add($p) } }
    }
    $result.pagesTotal = $total
    $result.pagesPrinted = $sel.label
    $result.order = Format-PageList @($order)
    $result.plan = 'kind=' + $result.kind + '; pages=' + $total + '; selected=' + $sel.label + '; copies=' + $Copies + '; collate=' + $collateText + '; duplex=' + $Duplex + '(' + $duplexName + '); order=' + $result.order
    if ($PlanOnly) { $result.ok = $true; return $result }

    $doc.PrinterSettings.Copies = 1
    $doc.PrinterSettings.Collate = $false
    $doc.PrinterSettings.Duplex = [System.Enum]::Parse([System.Drawing.Printing.Duplex], $duplexName)
    if ($OutputFile) {
      $out = [System.IO.Path]::GetFullPath($OutputFile)
      $outDir = [System.IO.Path]::GetDirectoryName($out)
      if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
      if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
      $doc.PrinterSettings.PrintToFile = $true
      $doc.PrinterSettings.PrintFileName = $out
      $result.output = $out
    }
    if ($isImage) { $image = [System.Drawing.Image]::FromFile($Path) }
    # The page sequence is decided here, so copies and their order never depend on whether the
    # driver honours DEVMODE copies/collate: page 1..n per copy = collated, n x page 1..n = stacked.
    $state = @{ order = @($order); index = 0; pages = $pages; image = $image; isImage = $isImage; font = $font; lineHeight = $lineHeight }
    $handler = {
      param($sender, $e)
      $i = [int]$state.index
      if ($i -ge $state.order.Count) { $e.HasMorePages = $false; return }
      $pageNo = [int]$state.order[$i]
      $state.index = $i + 1
      if ($state.isImage) {
        $null = $state.image.SelectActiveFrame([System.Drawing.Imaging.FrameDimension]::Page, ($pageNo - 1))
        $b = $e.MarginBounds
        $iw = [int]$state.image.Width
        $ih = [int]$state.image.Height
        if ($iw -gt 0 -and $ih -gt 0 -and $b.Width -gt 0 -and $b.Height -gt 0) {
          $scale = [Math]::Min(([double]$b.Width / $iw), ([double]$b.Height / $ih))
          $w = [int]($iw * $scale)
          $h = [int]($ih * $scale)
          if ($w -lt 1) { $w = 1 }
          if ($h -lt 1) { $h = 1 }
          $x = [int]$b.Left + [int](($b.Width - $w) / 2)
          $y = [int]$b.Top + [int](($b.Height - $h) / 2)
          $e.Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
          $rect = New-Object System.Drawing.Rectangle -ArgumentList $x, $y, $w, $h
          $e.Graphics.DrawImage($state.image, $rect)
        }
      } else {
        $bounds = $e.MarginBounds
        # Lines are wrapped to the printable width already; the clip is the last line of defence
        # against a glyph that is wider than measured, so nothing is drawn outside the margin box.
        $e.Graphics.SetClip($bounds)
        $x = [single]$bounds.Left
        $y = [single]$bounds.Top
        $limit = [single]$bounds.Bottom
        foreach ($line in $state.pages[$pageNo - 1]) {
          if (($y + $state.lineHeight) -gt $limit) { break }
          $e.Graphics.DrawString([string]$line, $state.font, [System.Drawing.Brushes]::Black, $x, $y)
          $y = $y + $state.lineHeight
        }
      }
      if ($state.index -ge $state.order.Count) { $e.HasMorePages = $false } else { $e.HasMorePages = $true }
    }.GetNewClosure()
    $doc.add_PrintPage($handler)
    $doc.Print()
    if ($result.output) {
      $size = -1
      for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 400
        if (Test-Path -LiteralPath $result.output) {
          $len = (Get-Item -LiteralPath $result.output).Length
          if ($len -gt 0 -and $len -eq $size) { break }
          $size = $len
        }
      }
      if (-not (Test-Path -LiteralPath $result.output) -or (Get-Item -LiteralPath $result.output).Length -le 0) {
        $result.error = 'the printer driver did not write ' + $result.output + ' (some drivers, Microsoft Print to PDF among them, ignore print-to-file - use mode=file for those)'
        $result.code = 'noOutput'
        return $result
      }
    }
    $result.ok = $true
    return $result
  } catch {
    $result.error = Gdi-ErrorText $_
    if (-not $result.code) { $result.code = 'printFailed' }
    return $result
  } finally {
    if ($image) { $image.Dispose() }
    if ($font) { $font.Dispose() }
    if ($doc) { $doc.Dispose() }
  }
}

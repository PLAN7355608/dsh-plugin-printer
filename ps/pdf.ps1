# PDF assembly for the scanner: multipage pdf that embeds each page jpeg as a DCTDecode image
# stream, so the bytes the scanner produced go in untouched (no re-encoding, no external tool).
# Dot-sourced by scan.ps1; test/pdf-check.ps1 exercises it on its own with generated jpegs.
#
# The pages are streamed to disk one at a time instead of being collected in memory, and the total
# size is capped (default 512 MB): a 50-page 600 dpi colour scan used to be held twice (object list
# plus final byte array) before anything was written. The cap is a safety net, not a target - a job
# that hits it is refused with a message that suggests scanning to jpg images instead.
if (-not $LF) { $LF = [string][char]10 }
if (-not $ASCII) { $ASCII = [System.Text.Encoding]::ASCII }

# Default ceiling for one generated pdf.
$PDF_MAX_BYTES = 536870912

# Reads width/height/component-count straight out of the jpeg header. Doing it here (instead of via
# System.Drawing) keeps the pdf writer independent of the decoder's guesses: /ColorSpace in the pdf
# must match how the decoder will treat the embedded DCT stream.
function Get-JpegInfo([string]$path) {
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $n = $bytes.Length
  $info = @{ width = 0; height = 0; components = 3; bytes = $bytes; length = $n; isJpeg = $false }
  if ($n -lt 4) { return $info }
  if (-not ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8)) { return $info }
  $end = -1
  for ($j = $n - 2; $j -ge 0; $j--) {
    if ($bytes[$j] -eq 0xFF -and $bytes[$j + 1] -eq 0xD9) { $end = $j + 2; break }
  }
  if ($end -lt 0) { return $info }
  $info.isJpeg = $true
  # Drop everything after the final EOI marker so the embedded stream is exactly one jpeg.
  $info.length = $end
  $i = 2
  while ($i -lt ($n - 9)) {
    if ($bytes[$i] -ne 0xFF) { $i++; continue }
    $marker = [int]$bytes[$i + 1]
    if ($marker -eq 0x01 -or ($marker -ge 0xD0 -and $marker -le 0xD9)) { $i += 2; continue }
    if ($marker -eq 0xDA) { break }
    $len = ([int]$bytes[$i + 2] * 256) + [int]$bytes[$i + 3]
    $isSof = (($marker -ge 0xC0 -and $marker -le 0xCF) -and $marker -ne 0xC4 -and $marker -ne 0xC8 -and $marker -ne 0xCC)
    if ($isSof) {
      $info.height = ([int]$bytes[$i + 5] * 256) + [int]$bytes[$i + 6]
      $info.width = ([int]$bytes[$i + 7] * 256) + [int]$bytes[$i + 8]
      $info.components = [int]$bytes[$i + 9]
      break
    }
    $i += 2 + $len
  }
  return $info
}

# Builds a multi-page pdf that embeds each page jpeg as a DCTDecode image stream. No external tool, no
# re-encoding: the bytes the scanner produced go in untouched, so a scan stays lossless and the pdf is
# only as large as the pages are. Every page is checked to be a real jpeg before anything is written.
function Write-Pdf([string[]]$images, [string]$output, [int]$dpi, [long]$MaxBytes = 0) {
  if ($MaxBytes -le 0) { $MaxBytes = $PDF_MAX_BYTES }
  $count = @($images).Count
  if ($count -le 0) { throw 'there is no page image to put into the pdf' }
  if ($dpi -le 0) { $dpi = 300 }

  # Geometry only: the pixel bytes are read one page at a time while writing, so a long scan does not
  # have to sit in memory twice.
  $geoms = @()
  for ($k = 0; $k -lt $count; $k++) {
    $g = Get-JpegInfo $images[$k]
    if (-not $g.isJpeg) { throw ('page ' + [string]($k + 1) + ' is not a JPEG image (no JPEG start/end marker): ' + $images[$k]) }
    if ($g.width -le 0 -or $g.height -le 0) { throw ('page ' + [string]($k + 1) + ' has no usable JPEG frame header: ' + $images[$k]) }
    $g.bytes = $null
    $geoms += $g
  }

  $kids = @()
  for ($k = 0; $k -lt $count; $k++) { $kids += ([string](5 + 3 * $k) + ' 0 R') }

  $fs = [System.IO.File]::Create($output)
  try {
    $offsets = New-Object System.Collections.ArrayList
    # Writes one object body or header chunk, refusing to grow the file past the cap.
    function Add-PdfBytes([byte[]]$bytes) {
      if (($fs.Position + $bytes.Length) -gt $MaxBytes) {
        throw ('the pdf would exceed the ' + [string]$MaxBytes + ' byte limit. Scan to jpg images instead (format=jpg), or lower the resolution.')
      }
      $fs.Write($bytes, 0, $bytes.Length)
    }

    Add-PdfBytes $ASCII.GetBytes('%PDF-1.4' + $LF)
    $binary = New-Object byte[] 5
    $binary[0] = 0x25; $binary[1] = 0xE2; $binary[2] = 0xE3; $binary[3] = 0xCF; $binary[4] = 0xD3
    Add-PdfBytes $binary
    Add-PdfBytes (New-Object byte[] 1)

    [void]$offsets.Add($fs.Position)
    Add-PdfBytes $ASCII.GetBytes('1 0 obj' + $LF + '<< /Type /Catalog /Pages 2 0 R >>' + $LF + 'endobj' + $LF)
    [void]$offsets.Add($fs.Position)
    Add-PdfBytes $ASCII.GetBytes('2 0 obj' + $LF + '<< /Type /Pages /Kids [' + [string]::Join(' ', $kids) + '] /Count ' + [string]$count + ' >>' + $LF + 'endobj' + $LF)

    for ($k = 0; $k -lt $count; $k++) {
      $info = $geoms[$k]
      $imgObj = 3 + 3 * $k
      $contentObj = 4 + 3 * $k
      $pageObj = 5 + 3 * $k
      $space = 'DeviceRGB'
      $decode = ''
      if ($info.components -eq 1) { $space = 'DeviceGray' }
      if ($info.components -eq 4) {
        $space = 'DeviceCMYK'
        # Adobe's CMYK convention is inverted relative to device values: without /Decode the page
        # comes out as a negative. (Review suggestion 2.)
        $decode = ' /Decode [1 0 1 0 1 0 1 0]'
      }
      $wpt = [int][Math]::Round([double]$info.width * 72.0 / [double]$dpi)
      $hpt = [int][Math]::Round([double]$info.height * 72.0 / [double]$dpi)
      if ($wpt -lt 1) { $wpt = 1 }
      if ($hpt -lt 1) { $hpt = 1 }

      $imgDict = $LF + '<< /Type /XObject /Subtype /Image /Width ' + [string]$info.width + ' /Height ' + [string]$info.height + ' /ColorSpace /' + $space + $decode + ' /BitsPerComponent 8 /Filter /DCTDecode /Length ' + [string]$info.length + ' >>' + $LF + 'stream' + $LF
      $pageBytes = [System.IO.File]::ReadAllBytes($images[$k])
      $payload = New-Object byte[] $info.length
      [System.Array]::Copy($pageBytes, 0, $payload, 0, $info.length)
      $pageBytes = $null
      [void]$offsets.Add($fs.Position)
      Add-PdfBytes $ASCII.GetBytes([string]$imgObj + ' 0 obj' + $imgDict)
      Add-PdfBytes $payload
      $payload = $null
      Add-PdfBytes $ASCII.GetBytes($LF + 'endstream' + $LF + 'endobj' + $LF)

      $content = 'q' + $LF + [string]$wpt + ' 0 0 ' + [string]$hpt + ' 0 0 cm' + $LF + '/Im0 Do' + $LF + 'Q' + $LF
      $contentBytes = $ASCII.GetBytes($content)
      [void]$offsets.Add($fs.Position)
      Add-PdfBytes $ASCII.GetBytes([string]$contentObj + ' 0 obj' + $LF + '<< /Length ' + [string]$contentBytes.Length + ' >>' + $LF + 'stream' + $LF + $content + 'endstream' + $LF + 'endobj' + $LF)

      $pageDict = [string]$pageObj + ' 0 obj' + $LF + '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + [string]$wpt + ' ' + [string]$hpt + '] /Resources << /XObject << /Im0 ' + [string]$imgObj + ' 0 R >> >> /Contents ' + [string]$contentObj + ' 0 R >>' + $LF + 'endobj' + $LF
      [void]$offsets.Add($fs.Position)
      Add-PdfBytes $ASCII.GetBytes($pageDict)
    }

    $xrefPos = $fs.Position
    # A cross-reference offset is exactly ten digits; a file that grew past that would produce a pdf
    # no reader can parse, so it is refused instead of being written wrong.
    foreach ($off in $offsets) {
      if ($off -gt 9999999999) { throw ('the pdf is too large to index: cross-reference offset ' + [string]$off + ' does not fit in ten digits') }
    }
    if ($xrefPos -gt 9999999999) { throw ('the pdf is too large to index: cross-reference offset ' + [string]$xrefPos + ' does not fit in ten digits') }
    $xref = 'xref' + $LF + '0 ' + [string]($offsets.Count + 1) + $LF + '0000000000 65535 f ' + $LF
    foreach ($off in $offsets) { $xref += ([string]$off).PadLeft(10, '0') + ' 00000 n ' + $LF }
    $trailer = 'trailer' + $LF + '<< /Size ' + [string]($offsets.Count + 1) + ' /Root 1 0 R >>' + $LF + 'startxref' + $LF + [string]$xrefPos + $LF + '%%EOF' + $LF
    Add-PdfBytes $ASCII.GetBytes($xref + $trailer)

    $total = $fs.Length
  } finally {
    $fs.Dispose()
  }
  return $total
}

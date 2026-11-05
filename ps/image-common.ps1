# Image helpers shared by scan.ps1 and the guard checks: format sniffing, re-encoding and the
# A4 crop of a page that was fed on a longer sheet. Kept in its own file so the cropping and
# resolution rules can be regression-tested without a scanner attached.
function Get-ImageFormat([string]$name) {
  switch ($name) {
    'jpg'  { return [System.Drawing.Imaging.ImageFormat]::Jpeg }
    'jpeg' { return [System.Drawing.Imaging.ImageFormat]::Jpeg }
    'png'  { return [System.Drawing.Imaging.ImageFormat]::Png }
    'tif'  { return [System.Drawing.Imaging.ImageFormat]::Tiff }
    'tiff' { return [System.Drawing.Imaging.ImageFormat]::Tiff }
    'bmp'  { return [System.Drawing.Imaging.ImageFormat]::Bmp }
    default { return $null }
  }
}


# Reads the file header, because a driver may hand back a different format than requested.
function Get-FileFormat([string]$path) {
  try {
    $fs = [System.IO.File]::OpenRead($path)
    try {
      $buf = New-Object byte[] 12
      [void]$fs.Read($buf, 0, 12)
    } finally { $fs.Dispose() }
    if ($buf[0] -eq 0x42 -and $buf[1] -eq 0x4D) { return 'bmp' }
    if ($buf[0] -eq 0xFF -and $buf[1] -eq 0xD8) { return 'jpg' }
    if ($buf[0] -eq 0x89 -and $buf[1] -eq 0x50) { return 'png' }
    if (($buf[0] -eq 0x49 -and $buf[1] -eq 0x49) -or ($buf[0] -eq 0x4D -and $buf[1] -eq 0x4D)) { return 'tif' }
  } catch { }
  return ''
}

# Re-encodes a page into the requested format (used when the driver ignored the request).
function Convert-ImageFile([string]$path, [string]$format, [int]$Dpi = 0) {
  Add-Type -AssemblyName System.Drawing -ErrorAction Stop
  $target = Get-ImageFormat $format
  if (-not $target) { throw ('cannot convert to ' + $format) }
  $img = [System.Drawing.Image]::FromFile($path)
  $tmp = $path + '.converting'
  try {
    if ($Dpi -gt 0) {
      # Re-encode through a bitmap so the saved page carries the real resolution: a driver that
      # reports 96 dpi in the header would otherwise leave every page mislabelled.
      $bmp = New-Object System.Drawing.Bitmap -ArgumentList @($img.Width, $img.Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
      try {
        $bmp.SetResolution($Dpi, $Dpi)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try { $g.DrawImage($img, 0, 0, $img.Width, $img.Height) } finally { $g.Dispose() }
        $bmp.Save($tmp, $target)
      } finally { $bmp.Dispose() }
    } else {
      $img.Save($tmp, $target)
    }
  } finally {
    $img.Dispose()
  }
  Remove-Item -LiteralPath $path -Force
  Move-Item -LiteralPath $tmp -Destination $path -Force
  return $true
}

function Get-ImageSize([string]$path) {
  $info = @{ width = 0; height = 0; dpi = 0 }
  try {
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $img = [System.Drawing.Image]::FromFile($path)
    try {
      $info.width = [int]$img.Width
      $info.height = [int]$img.Height
      $info.dpi = [int][Math]::Round([double]$img.HorizontalResolution)
    } finally { $img.Dispose() }
  } catch { }
  return $info
}


# Trims the right/bottom overflow of a page that was fed on a longer sheet: a 14 inch document
# feeder produces a 14 inch image, so an A4 original comes back with a strip of blank paper.
# Content is assumed to start at the top left, which is how the sheet enters the feeder.
# The cropping ruler is the negotiated resolution rather than the header one: drivers round the
# resolution, and some of them (the HP PCL-6 driver used for the tests) stamp a bogus 96 dpi
# into the file header while really scanning at 300 dpi.
function Crop-Image([string]$path, [int]$dpi, [string]$mode, [string]$format) {
  $res = @{ cropped = $false; dpi = $dpi; dpiMismatch = $false }
  if ($mode -ne 'a4') { return $res }
  Add-Type -AssemblyName System.Drawing -ErrorAction Stop
  $img = [System.Drawing.Image]::FromFile($path)
  try {
    $imgDpi = 0
    try { $imgDpi = [int][Math]::Round([double]$img.HorizontalResolution) } catch { $imgDpi = 0 }
    if ($dpi -gt 0 -and $imgDpi -gt 0 -and [Math]::Abs($imgDpi - $dpi) -gt 2) { $res.dpiMismatch = $true }
    # The negotiated value is the one the driver confirmed during the scan, so it is the better ruler:
    # this HP driver stamps a bogus 96 dpi into the file header while really scanning at 300 dpi, and
    # cropping with the header value shrank an A4 page to a quarter of its size.
    $eff = $dpi
    if ($eff -le 0) { $eff = $imgDpi }
    $res.dpi = $eff
    if ($eff -le 0) { return $res }
    $tw = [int][Math]::Round(8.2677 * $eff)
    $th = [int][Math]::Round(11.6929 * $eff)
    $w = [int]$img.Width
    $h = [int]$img.Height
    $nw = [Math]::Min($w, $tw)
    $nh = [Math]::Min($h, $th)
    if ($nw -eq $w -and $nh -eq $h) { return $res }
    $enc = Get-ImageFormat $format
    if (-not $enc) { $enc = [System.Drawing.Imaging.ImageFormat]::Jpeg }
    $bmp = New-Object -TypeName System.Drawing.Bitmap -ArgumentList @($nw, $nh)
    $tmp = $path + '.cropped'
    try {
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      try {
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $rect = New-Object -TypeName System.Drawing.Rectangle -ArgumentList @(0, 0, $nw, $nh)
        $g.DrawImage($img, $rect, $rect, [System.Drawing.GraphicsUnit]::Pixel)
      } finally { $g.Dispose() }
      $bmp.Save($tmp, $enc)
    } finally { $bmp.Dispose() }
  } finally { $img.Dispose() }
  Remove-Item -LiteralPath $path -Force
  Move-Item -LiteralPath $tmp -Destination $path -Force
  $res.cropped = $true
  return $res
}

# eSCL (Scanner Control Language / Mopria) client shared by the scanner tools.
#
# A network scanner that advertises eSCL over HTTP(S) can run a whole ADF stack as ONE job: the job
# is created once and every NextDocument call hands back the next page. That is the protocol path used
# for multi-page feeder scans; the WIA path stays as the fallback for USB-only devices.
#
# Protocol notes learned from a real HP Color LaserJet MFP M283cdw (gSOAP/2.7):
#   * ScannerStatus answers on both http and https, but POST /eSCL/ScanJobs goes to the HTTPS port.
#   * POST /eSCL/ScanJobs answers 400 when the request carries "Expect: 100-continue", which
#     HttpWebRequest adds by default. ServicePointManager.Expect100Continue has to be turned off.
#   * The job URL arrives in the Location header; pages are pulled from <job>/NextDocument.
#   * NextDocument answers 404 once the batch is finished (no more pages), so 404/410 is the normal
#     end of a feeder run, not an error.
#   * A job that is never collected leaves the device in "Processing" and every new job answers 503,
#     so a new job starts by deleting the leftovers.
#   * Scan region units are 1/300 inch: A4 is 2480 x 3508 regardless of the output resolution.

function Enable-EsclTls {
  # The scanner presents a self-signed certificate; the caller is a local tool, so accept it.
  [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
  try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch { }
  # Without this the gSOAP server answers 400 to every POST.
  [System.Net.ServicePointManager]::Expect100Continue = $false
}

function Invoke-EsclHttp {
  param(
    [string]$Method = 'GET',
    [Parameter(Mandatory = $true)][string]$Url,
    [byte[]]$Body,
    [string]$OutFile,
    [int]$TimeoutSec = 60
  )
  Enable-EsclTls
  $result = @{ ok = $false; status = 0; text = ''; bytes = 0; contentType = ''; location = ''; error = ''; gone = $false }
  $req = $null
  $resp = $null
  try {
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method = $Method
    $req.Timeout = $TimeoutSec * 1000
    $req.ReadWriteTimeout = $TimeoutSec * 1000
    $req.KeepAlive = $false
    $req.AllowAutoRedirect = $true
    $req.UserAgent = 'dsh-plugin-printer'
    $req.Accept = 'text/xml, image/jpeg, application/pdf, */*'
    if ($Body) {
      $req.ContentType = 'text/xml'
      $req.ContentLength = $Body.Length
      $stream = $req.GetRequestStream()
      try { $stream.Write($Body, 0, $Body.Length) } finally { $stream.Close() }
    }
    try {
      $resp = $req.GetResponse()
    } catch [System.Net.WebException] {
      $we = $_.Exception
      if ($we.Response) {
        $resp = $we.Response
        $result.status = [int]$resp.StatusCode
      } else {
        $result.error = $we.Message
      }
    }
    if ($resp) {
      $result.status = [int]$resp.StatusCode
      $result.contentType = [string]$resp.ContentType
      if ($resp.Headers['Location']) { $result.location = [string]$resp.Headers['Location'] }
      $responseStream = $resp.GetResponseStream()
      try {
        if ($OutFile) {
          $fileStream = [System.IO.File]::Create($OutFile)
          try { $responseStream.CopyTo($fileStream) } finally { $fileStream.Close() }
          $result.bytes = (Get-Item -LiteralPath $OutFile).Length
        } else {
          $reader = New-Object System.IO.StreamReader($responseStream, [System.Text.Encoding]::UTF8)
          try { $result.text = $reader.ReadToEnd() } finally { $reader.Close() }
        }
      } finally { $responseStream.Close() }
      $result.ok = ($result.status -ge 200 -and $result.status -lt 300)
      if ($result.status -eq 404 -or $result.status -eq 410) { $result.gone = $true }
    }
  } catch {
    $result.error = $_.Exception.Message
  } finally {
    if ($resp) { try { $resp.Close() } catch { } }
  }
  return $result
}

# Normalises whatever the caller knows about the device into one base URL that answers eSCL.
# Returns @{ ok; base; error }.
function Resolve-EsclBase {
  param([string]$Hint = '')
  $result = @{ ok = $false; base = ''; error = '' }
  $candidates = @()
  $override = $env:DSH_ESCL_BASE
  if ($override) { $candidates += $override.TrimEnd('/') }
  if ($Hint) {
    $h = $Hint.Trim()
    if ($h -match '^(https?)://') {
      try {
        # The WIA presentation URL usually carries port 80 (http). The scan jobs live on the HTTPS
        # port, so probe the bare host first and keep the advertised URL as the last resort.
        $uri = [System.Uri]$h
        $candidates += ('https://' + $uri.Host)
        $candidates += ('http://' + $uri.Host)
        $candidates += ($h.TrimEnd('/'))
      } catch {
        $candidates += ($h.TrimEnd('/'))
      }
    } elseif ($h -match '^[0-9a-zA-Z_.-]+(:[0-9]+)?$') {
      $candidates += ('https://' + $h)
      $candidates += ('http://' + $h)
    }
  }
  $seen = @{}
  foreach ($candidate in $candidates) {
    if (-not $candidate -or $seen.ContainsKey($candidate)) { continue }
    $seen[$candidate] = $true
    $probe = Invoke-EsclHttp -Url ($candidate + '/eSCL/ScannerStatus') -TimeoutSec 8
    if ($probe.ok -and $probe.text -match 'ScannerStatus') {
      $result.ok = $true
      $result.base = $candidate
      return $result
    }
    if (-not $result.error -and $probe.error) { $result.error = $probe.error }
  }
  if (-not $result.error) { $result.error = 'the device did not answer eSCL on its advertised address' }
  return $result
}

# Reads the address the WIA driver knows about (the WSD presentation URL) and resolves eSCL from it.
# The device has to be open before the driver publishes that property, so callers pass an open item.
function Get-EsclBaseFromWiaDevice {
  param($Info)
  $hint = ''
  if ($Info) {
    foreach ($id in 38919, 38920, 6, 3114) {
      $value = $null
      try { $value = Get-PropertyById $Info.Properties $id } catch { $value = $null }
      if ($value -and ([string]$value -match '^(https?)://')) { $hint = [string]$value; break }
    }
  }
  return (Resolve-EsclBase $hint)
}

# Returns @{ ok; state; adfState; error }.
function Get-EsclStatus {
  param([Parameter(Mandatory = $true)][string]$Base)
  $r = Invoke-EsclHttp -Url ($Base + '/eSCL/ScannerStatus') -TimeoutSec 10
  $out = @{ ok = $false; state = ''; adfState = ''; error = '' }
  if (-not $r.ok) { $out.error = $r.error; return $out }
  $out.ok = $true
  $m = [regex]::Match($r.text, '<pwg:State>([^<]*)</pwg:State>')
  if ($m.Success) { $out.state = $m.Groups[1].Value }
  $m = [regex]::Match($r.text, '<scan:AdfState>([^<]*)</scan:AdfState>')
  if ($m.Success) { $out.adfState = $m.Groups[1].Value }
  return $out
}

function Get-EsclResolutionList {
  param([string]$Text)
  $list = @()
  if (-not $Text) { return $list }
  foreach ($m in [regex]::Matches($Text, '<scan:XResolution>([0-9]+)</scan:XResolution>')) {
    $value = [int]$m.Groups[1].Value
    if ($list -notcontains $value) { $list += $value }
  }
  return $list
}

# Returns @{ ok; platenResolutions; adfResolutions; adfMaxWidth; adfMaxHeight; formats; error }.
function Get-EsclCapabilities {
  param([Parameter(Mandatory = $true)][string]$Base)
  $out = @{ ok = $false; platenResolutions = @(); adfResolutions = @(); adfMaxWidth = 0; adfMaxHeight = 0; formats = @(); error = '' }
  $r = Invoke-EsclHttp -Url ($Base + '/eSCL/ScannerCapabilities') -TimeoutSec 15
  if (-not $r.ok -and $r.error) {
    # Some firmware only answers capabilities over plain HTTP.
    $alt = $Base -replace '^https:', 'http:'
    $r = Invoke-EsclHttp -Url ($alt + '/eSCL/ScannerCapabilities') -TimeoutSec 15
  }
  if (-not $r.ok) { $out.error = $r.error; return $out }
  $text = $r.text
  $out.ok = $true
  # Split the document into the platen block and the ADF block so each keeps its own limits.
  $platenText = ''
  $adfText = ''
  $m = [regex]::Match($text, '(?s)<scan:Platen>.*?</scan:Platen>')
  if ($m.Success) { $platenText = $m.Value }
  $m = [regex]::Match($text, '(?s)<scan:Adf>.*?</scan:Adf>')
  if ($m.Success) { $adfText = $m.Value }
  $out.platenResolutions = @(Get-EsclResolutionList $platenText)
  $out.adfResolutions = @(Get-EsclResolutionList $adfText)
  foreach ($pair in @(@('adfMaxWidth', 'MaxWidth'), @('adfMaxHeight', 'MaxHeight'))) {
    $m = [regex]::Match($adfText, '<scan:' + $pair[1] + '>([0-9]+)</scan:' + $pair[1] + '>')
    if ($m.Success) { $out[$pair[0]] = [int]$m.Groups[1].Value }
  }
  $formats = @()
  foreach ($fm in [regex]::Matches($text, '<pwg:DocumentFormat>([^<]*)</pwg:DocumentFormat>')) {
    $value = $fm.Groups[1].Value
    if ($formats -notcontains $value) { $formats += $value }
  }
  $out.formats = $formats
  return $out
}

function Select-EsclResolution {
  param([int[]]$Accepted, [int]$Want)
  if (-not $Accepted -or $Accepted.Count -eq 0) { return $Want }
  $best = $Accepted[0]
  $bestDelta = [Math]::Abs($best - $Want)
  foreach ($value in $Accepted) {
    $delta = [Math]::Abs([int]$value - $Want)
    if ($delta -lt $bestDelta) { $best = [int]$value; $bestDelta = $delta }
  }
  return $best
}

# Builds the ScanSettings document. Region is in 1/300 inch, so A4 is always 2480 x 3508.
function New-EsclSettingsXml {
  param(
    [Parameter(Mandatory = $true)][string]$InputSource,
    [int]$Resolution = 300,
    [string]$ColorMode = 'RGB24',
    [string]$Intent = 'Document',
    [int]$Width = 2480,
    [int]$Height = 3508,
    [string]$DocumentFormat = 'image/jpeg'
  )
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8"?>')
  [void]$sb.Append('<scan:ScanSettings xmlns:scan="http://schemas.hp.com/imaging/escl/2011/05/03" xmlns:pwg="http://www.pwg.org/schemas/2010/12/sm">')
  [void]$sb.Append('<pwg:Version>2.63</pwg:Version>')
  [void]$sb.Append('<pwg:ScanRegions pwg:MustHonor="true"><pwg:ScanRegion><pwg:XOffset>0</pwg:XOffset><pwg:YOffset>0</pwg:YOffset><pwg:Width>' + $Width + '</pwg:Width><pwg:Height>' + $Height + '</pwg:Height><pwg:ContentRegionUnits>escl:ThreeHundredthsOfInches</pwg:ContentRegionUnits></pwg:ScanRegion></pwg:ScanRegions>')
  [void]$sb.Append('<pwg:InputSource>' + $InputSource + '</pwg:InputSource>')
  [void]$sb.Append('<scan:ColorMode>' + $ColorMode + '</scan:ColorMode>')
  [void]$sb.Append('<scan:Intent>' + $Intent + '</scan:Intent>')
  [void]$sb.Append('<scan:XResolution>' + $Resolution + '</scan:XResolution><scan:YResolution>' + $Resolution + '</scan:YResolution>')
  [void]$sb.Append('<pwg:DocumentFormat>' + $DocumentFormat + '</pwg:DocumentFormat>')
  [void]$sb.Append('</scan:ScanSettings>')
  return $sb.ToString()
}

function Remove-EsclJob {
  param([string]$JobUrl)
  if (-not $JobUrl) { return }
  Invoke-EsclHttp -Method 'DELETE' -Url $JobUrl -TimeoutSec 10 | Out-Null
}

# Deletes job URLs the scanner still lists. A job that was never collected keeps the device in
# "Processing" and makes every new job fail with 503, so a new job starts by clearing the leftovers.
function Clear-EsclStaleJobs {
  param([Parameter(Mandatory = $true)][string]$Base)
  $r = Invoke-EsclHttp -Url ($Base + '/eSCL/ScannerStatus') -TimeoutSec 10
  if (-not $r.ok) { return }
  foreach ($m in [regex]::Matches($r.text, '<pwg:JobUri>([^<]*)</pwg:JobUri>')) {
    $uri = [string]$m.Groups[1].Value
    if ($uri) {
      $jobUrl = $uri
      if ($jobUrl -notmatch '^(https?)://') { $jobUrl = $Base + '/' + $jobUrl.TrimStart('/') }
      Remove-EsclJob -JobUrl $jobUrl
    }
  }
}

# Waits until the scanner leaves "Processing" so a fresh job is accepted (max ~20 s).
function Wait-EsclIdle {
  param([Parameter(Mandatory = $true)][string]$Base, [int]$TimeoutSec = 20)
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  while ((Get-Date) -lt $deadline) {
    $status = Get-EsclStatus $Base
    if (-not $status.ok) { return }
    if ($status.state -ne 'Processing') { return }
    Start-Sleep -Milliseconds 800
  }
}

# Creates a scan job and returns @{ ok; jobUrl; status; error }.
function New-EsclJob {
  param([Parameter(Mandatory = $true)][string]$Base, [Parameter(Mandatory = $true)][string]$SettingsXml)
  $body = [System.Text.Encoding]::UTF8.GetBytes($SettingsXml)
  $out = @{ ok = $false; jobUrl = ''; status = 0; error = '' }
  $r = $null
  for ($attempt = 1; $attempt -le 8; $attempt++) {
    $r = Invoke-EsclHttp -Method 'POST' -Url ($Base + '/eSCL/ScanJobs') -Body $body -TimeoutSec 60
    if ($r.ok) { break }
    if ($r.status -ne 409 -and $r.status -ne 503) { break }
    if ($attempt -eq 1) {
      Clear-EsclStaleJobs $Base
      Wait-EsclIdle $Base
    } else {
      Start-Sleep -Milliseconds 1200
    }
  }
  $out.status = $r.status
  if (-not $r.ok) {
    $out.error = $r.error
    if (-not $out.error) { $out.error = ('the scanner refused the scan job (HTTP ' + [string]$r.status + ')') }
    return $out
  }
  if (-not $r.location) {
    $out.error = 'the scanner accepted the job but did not say where to collect it'
    return $out
  }
  $out.ok = $true
  $out.jobUrl = $r.location.TrimEnd('/')
  return $out
}

# Pulls the next page of the job into OutFile. gone = 1 means the batch is finished.
function Get-EsclDocument {
  param([Parameter(Mandatory = $true)][string]$JobUrl, [Parameter(Mandatory = $true)][string]$OutFile, [int]$TimeoutSec = 300)
  if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force }
  $r = Invoke-EsclHttp -Url ($JobUrl + '/NextDocument') -OutFile $OutFile -TimeoutSec $TimeoutSec
  $out = @{ ok = $r.ok; status = $r.status; bytes = $r.bytes; contentType = $r.contentType; gone = $r.gone; error = $r.error }
  if ($r.ok -and $r.bytes -le 0) {
    $out.ok = $false
    $out.error = 'the scanner sent an empty page'
  }
  return $out
}

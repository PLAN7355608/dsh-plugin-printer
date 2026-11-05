param(
  [Parameter(Mandatory = $true)][string]$Action,
  [string]$Printer = '',
  [string]$JobId = ''
)
# Print queue operations through Win32_PrintJob (works on Windows 7, where the PrintManagement
# cmdlets Get-Printer/Get-PrintJob do not exist).
#   -Action list    [-Printer <name>]      list queued jobs
#   -Action cancel  -JobId <id> [-Printer <name>]
#   -Action purge   [-Printer <name>]      cancel every queued job of that printer
#
# Deleting a job is only a REQUEST: the spooler can keep the job listed (status "Deleting") or refuse
# silently. So every cancel/purge reads the queue back and reports what the queue actually contains
# afterwards; "Delete() did not throw" is never reported as success. The Win32_PrintJob query needs a
# normal user token, so a read failure is reported as its own error, never as "the queue is empty".
#
# Output records (US = 0x1f):
#   JOB      <US> Printer <US> JobId <US> Document <US> Status <US> Pages <US> Owner
#   PROGRESS <US> purge <US> Printer <US> requested <US> deleted <US> failed
#   OK       <US> cancel <US> Printer <US> JobId <US> readback <US> status
#   OK       <US> purge  <US> Printer <US> requested <US> deleted <US> remaining <US> readback
#   WARN     <US> jobs   <US> name <US> message
#   CODE     <US> code   (always immediately followed by ERR)
#   ERR      <US> message
#   DONE
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

function Split-JobName([string]$name) {
  $printer = $name
  $id = ''
  $pos = $name.LastIndexOf(',')
  if ($pos -gt 0) {
    $printer = $name.Substring(0, $pos).Trim()
    $id = $name.Substring($pos + 1).Trim()
  }
  return @($printer, $id)
}

# One queue read. @{ ok; error; jobs } - ok=false means the queue could not be read at all.
function Get-QueueJobs {
  $info = @{ ok = $false; error = ''; jobs = @() }
  try {
    $info.jobs = @(Get-WmiObject -Class Win32_PrintJob -ErrorAction Stop)
    $info.ok = $true
  } catch {
    $info.error = Get-ErrorText $_
  }
  return $info
}

# Job ids currently listed, optionally for one printer. Unparsable entries are reported, not silently
# dropped: a job we cannot name is not the same as no job.
function Get-QueueJobIds([string]$printer) {
  $info = @{ ok = $false; error = ''; ids = @() }
  $snap = Get-QueueJobs
  if (-not $snap.ok) { $info.error = $snap.error; return $info }
  $ids = @()
  foreach ($j in $snap.jobs) {
    $parts = Split-JobName ([string]$j.Name)
    if (-not $parts[1]) {
      Emit @('WARN', 'jobs', [string]$j.Name, 'this queue entry has no job number and cannot be matched to an id')
      continue
    }
    if ($printer -and $parts[0] -ne $printer) { continue }
    $ids += $parts[1]
  }
  $info.ok = $true
  $info.ids = $ids
  return $info
}

# The spooler may list a deleted job for a moment, so "still there" is only decided after polling.
function Wait-JobGone([string]$printer, [string]$id, [int]$timeoutMs) {
  $deadline = (Get-Date).AddMilliseconds($timeoutMs)
  $lastStatus = ''
  while ($true) {
    $now = Get-QueueJobs
    if (-not $now.ok) { return @{ ok = $false; error = $now.error; gone = $false; status = '' } }
    $found = $null
    foreach ($j in $now.jobs) {
      $parts = Split-JobName ([string]$j.Name)
      if ($parts[1] -eq $id -and (-not $printer -or $parts[0] -eq $printer)) { $found = $j; break }
    }
    if (-not $found) { return @{ ok = $true; error = ''; gone = $true; status = '' } }
    try { $lastStatus = [string]$found.JobStatus } catch { $lastStatus = '' }
    if ((Get-Date) -ge $deadline) { return @{ ok = $true; error = ''; gone = $false; status = $lastStatus } }
    Start-Sleep -Milliseconds 350
  }
}

try {
  $snap = Get-QueueJobs
  if (-not $snap.ok) {
    Fail ('cannot read the print queue (Win32_PrintJob query failed: ' + $snap.error + '). This is different from an empty queue: the spooler refused the query, usually because the process lacks the required privilege or the Print Spooler service is not running.') 'queueUnreadable'
  }

  switch ($Action) {
    'list' {
      foreach ($j in $snap.jobs) {
        $parts = Split-JobName ([string]$j.Name)
        if (-not $parts[1]) {
          Emit @('WARN', 'jobs', [string]$j.Name, 'this queue entry has no job number and is listed without an id')
        }
        if ($Printer -and $parts[0] -ne $Printer) { continue }
        Emit @('JOB', $parts[0], $parts[1], [string]$j.Document, [string]$j.JobStatus, [string]$j.TotalPages, [string]$j.Owner)
      }
      Emit @('DONE')
    }
    'cancel' {
      if (-not $JobId) { Fail 'no job id given' 'noJobId' }
      $targets = @()
      foreach ($j in $snap.jobs) {
        $parts = Split-JobName ([string]$j.Name)
        if ($parts[1] -ne $JobId) { continue }
        if ($Printer -and $parts[0] -ne $Printer) { continue }
        $targets += $j
      }
      if ($targets.Count -eq 0) { Fail ('job not found: ' + $JobId) 'jobNotFound' }
      $deleted = 0
      $failures = @()
      foreach ($j in $targets) {
        try { $j.Delete(); $deleted++ } catch { $failures += (Get-ErrorText $_) }
      }
      if ($deleted -eq 0) { Fail ('the spooler refused to delete job ' + $JobId + ': ' + ($failures -join '; ')) 'deleteFailed' }
      $gone = Wait-JobGone $Printer $JobId 4000
      if (-not $gone.ok) {
        Fail ('job ' + $JobId + ' was asked to delete, but the queue could not be read back to verify it: ' + $gone.error) 'verifyFailed'
      }
      if (-not $gone.gone) {
        Emit @('OK', 'cancel', $Printer, $JobId, 'present', $gone.status)
        Fail ('job ' + $JobId + ' is still in the queue after the delete request (status: ' + $gone.status + '). The spooler can hold a job in this state until the printer accepts the deletion; check the queue again.') 'stillQueued'
      }
      Emit @('OK', 'cancel', $Printer, $JobId, 'gone', '')
    }
    'purge' {
      $before = Get-QueueJobIds $Printer
      $requested = @($before.ids).Count
      $deleted = 0
      $failures = @()
      foreach ($j in $snap.jobs) {
        $parts = Split-JobName ([string]$j.Name)
        if ($Printer -and $parts[0] -ne $Printer) { continue }
        try { $j.Delete(); $deleted++ } catch { $failures += (Split-JobName ([string]$j.Name))[1] + ': ' + (Get-ErrorText $_) }
      }
      # Report progress before any failure can end the run, so the host can say what did get deleted.
      Emit @('PROGRESS', 'purge', $Printer, [string]$requested, [string]$deleted, [string]$failures.Count)
      $after = Get-QueueJobIds $Printer
      if (-not $after.ok) {
        Fail ('the purge was sent, but the queue could not be read back to verify it: ' + $after.error) 'verifyFailed'
      }
      $remaining = $after.ids.Count
      if ($remaining -gt 0) {
        Emit @('OK', 'purge', $Printer, [string]$requested, [string]$deleted, [string]$remaining, 'present')
        Fail ('the purge was sent, but ' + $remaining + ' job(s) are still listed in the queue after the readback. The spooler can hold jobs in a deleting state until the printer accepts the deletion; check the queue again.') 'stillQueued'
      }
      Emit @('OK', 'purge', $Printer, [string]$requested, [string]$deleted, '0', 'gone')
    }
    default { Fail ('unknown action: ' + $Action) 'unknownAction' }
  }
} catch {
  Emit @('ERR', (Get-ErrorText $_))
  exit 1
}

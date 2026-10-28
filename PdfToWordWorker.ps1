<#
.SYNOPSIS
    Long-lived PDF-to-Word conversion host.

.DESCRIPTION
    Owns exactly one Microsoft Word instance for its whole lifetime and converts one
    document per request. Speaks newline-delimited JSON on stdin/stdout:

        -> {"id":"1","cmd":"convert","path":"C:\\in\\a.pdf","out":"C:\\out\\a.docx",
            "format":"docx","force":true,"timeout":120}
        <- {"id":"1","ok":true,"result":{...}}
        <- {"id":"1","ok":false,"error":{"type":"validation","message":"..."}}

    Commands: convert, ping, stats, recycle, shutdown.

    The worker always CREATES its Word instance (COM activation), never attaches to an
    existing one. The process id is recorded at startup, so the watchdog can terminate
    this worker's Word without ever touching another one.

    stdout is reserved for JSON. Every log line goes to stderr.
#>
[CmdletBinding()]
param(
    [ValidateRange(0, 100000)]
    [int]$RecycleAfter = 50,

    [ValidateRange(0, 86400)]
    [int]$DefaultTimeoutSeconds = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'PdfToWord.Core.ps1')

# UTF-8 on both pipes, no BOM. The JSON reader on the Python side decodes UTF-8 and
# would choke on a BOM at the start of the first line.
$utf8   = New-Object System.Text.UTF8Encoding($false)
$stdin  = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8, $false)
$stdout = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$stderr = New-Object System.IO.StreamWriter([Console]::OpenStandardError(), $utf8)
$stdout.AutoFlush = $true
$stderr.AutoFlush = $true
try { [Console]::InputEncoding  = $utf8 } catch { }
try { [Console]::OutputEncoding = $utf8 } catch { }

$script:Session         = $null
$script:JobsSinceRecycle = 0
$script:TotalJobs       = 0
$script:FailedJobs      = 0

function Write-WorkerLog {
    param([string]$Message)
    try {
        $stderr.WriteLine(("[{0:HH:mm:ss}] [worker {1}] {2}" -f (Get-Date), $PID, $Message))
    }
    catch { }
}

function Send-WorkerResponse {
    param($Payload)
    $json = $Payload | ConvertTo-Json -Compress -Depth 8
    $stdout.WriteLine($json)
}

<#
    ConvertFrom-Json yields a PSCustomObject; reading a missing property under
    Set-StrictMode throws, so every optional field goes through this helper.
#>
function Get-JsonValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    if ($null -eq $property.Value) { return $Default }
    return $property.Value
}

function Ensure-WordSession {
    if ($script:Session -and $script:Session.Word) {
        $wordPid = $script:Session.Pid
        if (-not $wordPid -or (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)) {
            return $script:Session
        }

        # Word died between two jobs (crash, antivirus, manual kill). Reusing the
        # stale COM object would turn the next conversion into an RPC failure, so
        # replace it before touching the document.
        Write-WorkerLog "Word pid $wordPid disappeared between jobs; replacing it."
        Recycle-WordSession -DelaySeconds 10
    }

    $script:Session = New-WordSession
    if (-not $script:Session.Pid) {
        throw [System.Exception]::new('Word started but its process id could not be determined; refusing to run without a safe watchdog.')
    }

    $script:JobsSinceRecycle = 0
    Write-WorkerLog "Word ready (pid $($script:Session.Pid), version $($script:Session.Version))."
    # Keep the client's record of the owned pid current across recycles. During
    # startup this arrives just before the "ready" message.
    Send-WorkerResponse ([ordered]@{
        type         = 'word'
        word_pid     = $script:Session.Pid
        word_version = $script:Session.Version
    })
    return $script:Session
}

<#
    Quits the current Word and waits before the next one is created. A Word killed by
    the watchdog needs a longer pause: repeatedly force-killing Word leaves the COM
    server in a state where the next activation stalls.
#>
function Recycle-WordSession {
    param([int]$DelaySeconds = 0)

    if ($script:Session) {
        $wordPid = $script:Session.Pid
        $wasAlive = $false
        if ($wordPid) {
            $wasAlive = [bool](Get-Process -Id $wordPid -ErrorAction SilentlyContinue)
        }

        try { Close-WordSession -Session $script:Session }
        catch { Write-WorkerLog "Recycle failed: $($_.Exception.Message)" }

        $script:Session = $null
        $script:JobsSinceRecycle = 0

        # Tell the client the old Word instance is gone before any cooling-off sleep.
        Send-WorkerResponse ([ordered]@{ type = 'word'; word_pid = $null; word_version = $null })

        if ($wordPid -and -not $wasAlive -and $DelaySeconds -lt 10) {
            $DelaySeconds = 10
        }
        if ($DelaySeconds -gt 0) { Start-Sleep -Seconds $DelaySeconds }
    }
    else {
        $script:JobsSinceRecycle = 0
    }
}

<#
    Runs one convert request. Validation happens before Word is touched, so a bad path
    never costs a Word restart.
#>
function Invoke-WorkerJob {
    param($Request)

    $path         = [string](Get-JsonValue $Request 'path' '')
    $out          = [string](Get-JsonValue $Request 'out' '')
    $format       = [string](Get-JsonValue $Request 'format' 'docx')
    $imagesFolder = [string](Get-JsonValue $Request 'images_folder' $script:PdfToWordDefaultImagesFolder)
    $imageFormat  = [string](Get-JsonValue $Request 'image_format' 'png')
    $force        = [bool](Get-JsonValue $Request 'force' $false)
    $timeout      = [int](Get-JsonValue $Request 'timeout' $DefaultTimeoutSeconds)

    if (-not $path) { throw [System.ArgumentException]::new('Missing "path".') }
    if (-not $out)  { throw [System.ArgumentException]::new('Missing "out".') }
    if ($imageFormat -notin @('png', 'original')) {
        throw [System.ArgumentException]::new("Unknown image_format '$imageFormat'. Use png or original.")
    }

    $jobStart = Get-Date

    # Resolve first: ArgumentException from here is the caller's fault, not Word's.
    $job = Resolve-PdfToWordJob -Path $path -OutPath $out -Format $format -Force:$force

    $session = Ensure-WordSession

    try {
        $info = Invoke-WordDocumentConversion -Word $session.Word -Spec $job.Spec -PdfPath $job.Pdf `
                    -OutFile $job.OutFile -AllowPng:($imageFormat -eq 'png') `
                    -WatchdogSeconds $timeout -WatchdogPid ([int]$session.Pid)

        Complete-PdfToWordOutput -Spec $job.Spec -OutputPath $job.OutFile `
            -ImagesFolderName $imagesFolder -ImageFormat $imageFormat -Since $jobStart -Force:$force
    }
    catch {
        # A document that failed validation or post-processing is not a usable output.
        # Do not delete a file that existed before this request (force was used); only
        # remove the partial result this attempt created.
        if (-not $job.OutExisted -and (Test-Path -LiteralPath $job.OutFile)) {
            Remove-Item -LiteralPath $job.OutFile -Force -ErrorAction SilentlyContinue
        }
        throw
    }

    $companions = Get-PdfToWordCompanions -Spec $job.Spec -OutputPath $job.OutFile `
        -SourcePath $job.Pdf -ImagesFolderName $imagesFolder -Since $jobStart

    return [pscustomobject]@{
        source          = $job.Pdf
        output          = $job.OutFile
        format          = $format
        pages           = $info.Pages
        size_kb         = [math]::Round((Get-Item -LiteralPath $job.OutFile).Length / 1KB, 1)
        seconds         = [math]::Round(((Get-Date) - $jobStart).TotalSeconds, 2)
        attempts        = 1
        companion_files = @($companions)
        pdf_warning_set = [bool]$session.PdfWarningSet
    }
}

# ---------- startup ----------
try {
    Ensure-WordSession | Out-Null
    Send-WorkerResponse ([ordered]@{
        type            = 'ready'
        pid             = $PID
        word_pid        = $script:Session.Pid
        word_version    = $script:Session.Version
        recycle_after   = $RecycleAfter
        default_timeout = $DefaultTimeoutSeconds
    })
}
catch {
    Write-WorkerLog "Startup failed: $($_.Exception.Message)"
    Send-WorkerResponse ([ordered]@{ type = 'error'; message = $_.Exception.Message })
    exit 1
}

Write-WorkerLog "Ready. Recycling Word every $RecycleAfter job(s); default timeout $DefaultTimeoutSeconds s."

# ---------- request loop ----------
try {
    while ($true) {
        $line = $null
        try { $line = $stdin.ReadLine() }
        catch { Write-WorkerLog "stdin read failed: $($_.Exception.Message)"; break }

        if ($null -eq $line) {
            Write-WorkerLog 'stdin closed; shutting down.'
            break
        }
        if (-not $line.Trim()) { continue }

        $request = $null
        try { $request = $line | ConvertFrom-Json }
        catch {
            Send-WorkerResponse ([ordered]@{
                id    = $null
                ok    = $false
                error = [ordered]@{ type = 'protocol'; message = "Invalid JSON: $($_.Exception.Message)" }
            })
            continue
        }

        $id  = Get-JsonValue $request 'id' $null
        $cmd = [string](Get-JsonValue $request 'cmd' '')

        switch ($cmd) {
            'ping' {
                $wordPid = $null
                $wordVersion = $null
                if ($script:Session) {
                    $wordPid = $script:Session.Pid
                    $wordVersion = $script:Session.Version
                }
                Send-WorkerResponse ([ordered]@{
                    id     = $id
                    ok     = $true
                    result = [ordered]@{
                        pid           = $PID
                        word_pid      = $wordPid
                        word_version  = $wordVersion
                        jobs          = $script:TotalJobs
                        failed        = $script:FailedJobs
                        since_recycle = $script:JobsSinceRecycle
                    }
                })
            }

            'stats' {
                $wordPid = $null
                if ($script:Session) { $wordPid = $script:Session.Pid }
                Send-WorkerResponse ([ordered]@{
                    id     = $id
                    ok     = $true
                    result = [ordered]@{
                        pid           = $PID
                        word_pid      = $wordPid
                        jobs          = $script:TotalJobs
                        failed        = $script:FailedJobs
                        since_recycle = $script:JobsSinceRecycle
                        recycle_after = $RecycleAfter
                    }
                })
            }

            'convert' {
                try {
                    $result = Invoke-WorkerJob -Request $request
                    $script:TotalJobs++
                    $script:JobsSinceRecycle++
                    Send-WorkerResponse ([ordered]@{ id = $id; ok = $true; result = $result })

                    if ($RecycleAfter -gt 0 -and $script:JobsSinceRecycle -ge $RecycleAfter) {
                        Write-WorkerLog "Recycling Word after $($script:JobsSinceRecycle) job(s)."
                        Recycle-WordSession -DelaySeconds 2
                        try { Ensure-WordSession | Out-Null }
                        catch { Write-WorkerLog "Word restart failed: $($_.Exception.Message)" }
                    }
                }
                catch {
                    $script:FailedJobs++
                    $message = $_.Exception.Message

                    if ($_.Exception -is [System.ArgumentException]) {
                        # Bad request: Word was never touched, so do not recycle it.
                        Write-WorkerLog "Validation error: $message"
                        Send-WorkerResponse ([ordered]@{
                            id    = $id
                            ok    = $false
                            error = [ordered]@{ type = 'validation'; message = $message }
                        })
                    }
                    else {
                        $type = 'word'
                        if ($_.Exception -is [System.TimeoutException]) { $type = 'timeout' }
                        Write-WorkerLog "Job failed ($type): $message"
                        Send-WorkerResponse ([ordered]@{
                            id    = $id
                            ok    = $false
                            error = [ordered]@{ type = $type; message = $message }
                        })

                        # A Word that just failed is never reused. If the watchdog killed
                        # it, Recycle-WordSession adds the 10 s cooling-off period. The
                        # replacement is created lazily by the next request, so a batch
                        # that ends on a failure does not start (and immediately quit) a
                        # fresh Word.
                        Recycle-WordSession -DelaySeconds 3
                    }
                }
            }

            'recycle' {
                try {
                    Recycle-WordSession -DelaySeconds 2
                    Ensure-WordSession | Out-Null
                    Send-WorkerResponse ([ordered]@{
                        id     = $id
                        ok     = $true
                        result = [ordered]@{
                            word_pid     = $script:Session.Pid
                            word_version = $script:Session.Version
                        }
                    })
                }
                catch {
                    Send-WorkerResponse ([ordered]@{
                        id    = $id
                        ok    = $false
                        error = [ordered]@{ type = 'word'; message = $_.Exception.Message }
                    })
                }
            }

            'shutdown' {
                $jobs = $script:TotalJobs
                $failed = $script:FailedJobs
                if ($script:Session) {
                    try { Close-WordSession -Session $script:Session }
                    catch { Write-WorkerLog "Shutdown: $($_.Exception.Message)" }
                    $script:Session = $null
                }
                Send-WorkerResponse ([ordered]@{
                    id     = $id
                    ok     = $true
                    result = [ordered]@{ jobs = $jobs; failed = $failed }
                })
                exit 0
            }

            default {
                Send-WorkerResponse ([ordered]@{
                    id    = $id
                    ok    = $false
                    error = [ordered]@{ type = 'protocol'; message = "Unknown command '$cmd'." }
                })
            }
        }
    }
}
finally {
    if ($script:Session) {
        Write-WorkerLog 'Closing Word.'
        try { Close-WordSession -Session $script:Session } catch { }
        $script:Session = $null
    }
}

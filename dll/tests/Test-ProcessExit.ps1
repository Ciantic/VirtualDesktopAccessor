#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$LibraryPath,
    [string]$AutoHotkeyPath = (Join-Path $env:ProgramFiles 'AutoHotkey\v2\AutoHotkey64.exe'),
    [string]$DebuggerPath,
    [ValidateRange(1, 100)][int]$Iterations = 3,
    [ValidateRange(1, 120)][int]$TimeoutSeconds = 30,
    [string]$OutputDirectory = (Join-Path $env:TEMP ('vda-process-exit-' + [guid]::NewGuid().ToString('N')))
)

$ErrorActionPreference = 'Stop'
if (!$LibraryPath) { $LibraryPath = Join-Path $PSScriptRoot '..\..\target\release\VirtualDesktopAccessor.dll' }
if (!$DebuggerPath) { $DebuggerPath = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\Debuggers\x64\cdb.exe' }
$LibraryPath = (Resolve-Path -LiteralPath $LibraryPath).Path
$AutoHotkeyPath = (Resolve-Path -LiteralPath $AutoHotkeyPath).Path
$DebuggerPath = (Resolve-Path -LiteralPath $DebuggerPath).Path
$probe = Join-Path $PSScriptRoot 'process-exit.ahk'
$null = New-Item -ItemType Directory -Path $OutputDirectory -Force
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$commands = Join-Path $OutputDirectory 'debugger-commands.txt'
@'
.sympath .
sxe -c ".echo VDA_ACCESS_VIOLATION; .exr -1; k; q" av
sxe -c ".echo VDA_PROCESS_EXIT; .lastevent; q" epr
g
'@ | Set-Content -LiteralPath $commands -Encoding ASCII
$results = @()

# Each iteration gets a fresh process. An in-process test cannot assert that
# native TLS destructors survive ExitProcess. Never attach to a user's AHK.
foreach ($mode in @('load-only', 'count')) {
    for ($iteration = 1; $iteration -le $Iterations; $iteration++) {
        $start = New-Object System.Diagnostics.ProcessStartInfo
        $start.FileName = $DebuggerPath
        # -o keeps any child processes in this isolated debugging session.
        $start.Arguments = '-o -cf "' + $commands + '" "' + $AutoHotkeyPath + '" /ErrorStdOut "' + $probe + '" "' + $LibraryPath + '" ' + $mode
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        $process = [System.Diagnostics.Process]::Start($start)
        $processId = $process.Id
        $outTask = $process.StandardOutput.ReadToEndAsync()
        $errTask = $process.StandardError.ReadToEndAsync()
        $timedOut = !$process.WaitForExit($TimeoutSeconds * 1000)
        if ($timedOut) {
            # This handle belongs only to the child created above.
            $process.Kill()
            if (!$process.WaitForExit(5000)) { throw 'Timed-out probe could not be stopped' }
        }
        $stdout = $outTask.Result
        $stderr = $errTask.Result
        $exitCode = $process.ExitCode
        $process.Dispose()
        $exitHex = '0x' + [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$exitCode), 0).ToString('X8')
        $ready = $stdout -match '(?m)^VDA_EXIT_READY\r?$'
        $queried = $mode -eq 'load-only' -or $stdout -match '(?m)^desktop_count=[1-9][0-9]*\r?$'
        $accessViolation = $stdout -match '(?m)^VDA_ACCESS_VIOLATION\r?$'
        $normalExit = $stdout -match '(?m)^VDA_PROCESS_EXIT\r?$' -and $stdout -match '(?m)^Last event: .*Exit process .*?, code 0(?: \(0\))?\r?$'
        $passed = !$timedOut -and $exitCode -eq 0 -and $ready -and $queried -and $normalExit -and !$accessViolation -and !$stderr
        $results += [pscustomobject]@{
            Mode=$mode;Iteration=$iteration;DebuggerProcessId=$processId;Passed=$passed
            DebuggerExitCode=$exitCode;DebuggerExitHex=$exitHex;TimedOut=$timedOut
            ReadyToExit=$ready;AccessViolation=$accessViolation;NormalExit=$normalExit
        }
        $stdout | Set-Content -LiteralPath (Join-Path $OutputDirectory "$mode-$iteration.stdout.txt") -Encoding UTF8
        $stderr | Set-Content -LiteralPath (Join-Path $OutputDirectory "$mode-$iteration.stderr.txt") -Encoding UTF8
    }
}
$results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'results.json') -Encoding UTF8
$results | Format-Table Mode,Iteration,Passed,AccessViolation,NormalExit,TimedOut -AutoSize
Write-Output "Probe logs: $OutputDirectory"
if (@($results | Where-Object { !$_.Passed }).Count) {
    throw 'Process-exit regression failed; inspect child exit codes and logs above'
}
Write-Output "PASS: all $($results.Count) child processes exited normally."

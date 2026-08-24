param(
    [Parameter(Mandatory = $true)][string]$Command,
    [string[]]$CommandArgs = @()
)

# Drives a stdio MCP server through initialize + tools/list and prints the tool names.
# Used as a release smoke test: it proves the published package starts and advertises
# its tools. It deliberately calls no tool, so it needs no Word installation.

$ErrorActionPreference = 'Stop'

$psi = [System.Diagnostics.ProcessStartInfo]::new()
$psi.FileName = $Command
foreach ($a in $CommandArgs) { $psi.ArgumentList.Add($a) }
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false

$proc = [System.Diagnostics.Process]::Start($psi)

function Send-Message($obj) {
    $json = $obj | ConvertTo-Json -Depth 10 -Compress
    $proc.StandardInput.WriteLine($json)
    $proc.StandardInput.Flush()
}

function Receive-Message([int]$id, [int]$timeoutSeconds = 60) {
    $deadline = (Get-Date).AddSeconds($timeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $readTask = $proc.StandardOutput.ReadLineAsync()
        if (-not $readTask.Wait([int]($deadline - (Get-Date)).TotalMilliseconds)) { break }

        $line = $readTask.Result
        if ($null -eq $line) { throw "The server closed stdout before answering request $id." }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $message = $null
        try { $message = $line | ConvertFrom-Json } catch { continue }
        if ($message.id -eq $id) { return $message }
    }
    throw "No answer to request $id within $timeoutSeconds seconds."
}

try {
    Send-Message @{
        jsonrpc = '2.0'
        id      = 1
        method  = 'initialize'
        params  = @{
            protocolVersion = '2024-11-05'
            capabilities    = @{}
            clientInfo      = @{ name = 'smoke-test'; version = '1.0.0' }
        }
    }

    $init = Receive-Message -id 1
    if ($init.error) { throw "initialize failed: $($init.error | ConvertTo-Json -Compress)" }
    Write-Host "initialize: $($init.result.serverInfo.name) $($init.result.serverInfo.version)"

    Send-Message @{ jsonrpc = '2.0'; method = 'notifications/initialized' }

    Send-Message @{ jsonrpc = '2.0'; id = 2; method = 'tools/list' }
    $list = Receive-Message -id 2
    if ($list.error) { throw "tools/list failed: $($list.error | ConvertTo-Json -Compress)" }

    $tools = @($list.result.tools | ForEach-Object { $_.name } | Sort-Object)
    if ($tools.Count -eq 0) { throw "The server advertises no tools." }

    Write-Host "tools ($($tools.Count)): $($tools -join ', ')"
    exit 0
}
finally {
    if (-not $proc.HasExited) {
        try { $proc.StandardInput.Close() } catch { }
        if (-not $proc.WaitForExit(5000)) { $proc.Kill($true) }
    }
    $stderr = $proc.StandardError.ReadToEnd()
    if ($stderr) { Write-Host "--- server stderr ---`n$stderr" }
}

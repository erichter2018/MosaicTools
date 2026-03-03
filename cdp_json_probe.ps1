# Quick probe: get editor.getJSON() from Mosaic's ProseMirror editors
# Must run with MT killed (only one CDP connection allowed)

$portFile = Get-ChildItem "$env:LOCALAPPDATA\Packages\MosaicInfoHub_*\LocalState\EBWebView\DevToolsActivePort" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $portFile) { Write-Host "No DevToolsActivePort file found"; exit 1 }
$port = (Get-Content $portFile.FullName)[0].Trim()
Write-Host "CDP port: $port"

# Find iframe target
$targets = Invoke-RestMethod "http://localhost:$port/json" -TimeoutSec 5
$iframe = $targets | Where-Object { $_.url -match 'rp\.radpair\.com.*reports' } | Select-Object -First 1
if (-not $iframe) { Write-Host "No iframe target found (is a study open?)"; exit 1 }
Write-Host "Iframe: $($iframe.url)"

# Connect WebSocket
$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ws.ConnectAsync($iframe.webSocketDebuggerUrl, [System.Threading.CancellationToken]::None).Wait()

function Send-CDP($ws, $id, $js, $await=$false) {
    $p = @{ expression = $js; returnByValue = $true }
    if ($await) { $p.awaitPromise = $true }
    $msg = @{ id = $id; method = "Runtime.evaluate"; params = $p } | ConvertTo-Json -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($msg)
    $ws.SendAsync([ArraySegment[byte]]::new($bytes), 'Text', $true, [System.Threading.CancellationToken]::None).Wait()

    $buf = New-Object byte[] (4 * 1024 * 1024)  # 4MB buffer for large JSON
    $ms = New-Object System.IO.MemoryStream
    do {
        $result = $ws.ReceiveAsync([ArraySegment[byte]]::new($buf), [System.Threading.CancellationToken]::None).Result
        $ms.Write($buf, 0, $result.Count)
    } while (-not $result.EndOfMessage)

    $text = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
    $resp = $text | ConvertFrom-Json
    return $resp.result.result.value
}

# Get JSON from both editors
Write-Host "`n=== EDITOR 0 (Transcript) getJSON() ==="
$json0 = Send-CDP $ws 1 @"
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    if (!editors[0]?.editor) return 'no_editor_0';
    return JSON.stringify(editors[0].editor.getJSON(), null, 2);
})()
"@
Write-Host $json0

Write-Host "`n=== EDITOR 1 (Final Report) getJSON() ==="
$json1 = Send-CDP $ws 2 @"
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    if (!editors[1]?.editor) return 'no_editor_1';
    return JSON.stringify(editors[1].editor.getJSON(), null, 2);
})()
"@
Write-Host $json1

# Also get getHTML() for comparison
Write-Host "`n=== EDITOR 1 (Final Report) getHTML() ==="
$html1 = Send-CDP $ws 3 @"
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    if (!editors[1]?.editor) return 'no_editor_1';
    return editors[1].editor.getHTML();
})()
"@
Write-Host $html1

$ws.CloseAsync('NormalClosure', '', [System.Threading.CancellationToken]::None).Wait()
Write-Host "`nDone."

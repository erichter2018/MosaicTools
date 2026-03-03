# Quick CDP probe: test highlighting text in Mosaic's Tiptap editor
# Usage: powershell -ExecutionPolicy Bypass -File cdp_highlight_test.ps1

Add-Type -AssemblyName System.Net.WebSockets

$slimHubPort = $null
$pipes = [System.IO.Directory]::GetFiles("\\.\pipe\") | Where-Object { $_ -match "MosaicToolsCDP" }
if ($pipes) {
    $pipeClient = New-Object System.IO.Pipes.NamedPipeClientStream(".", ($pipes[0] -replace '.*\\',''), [System.IO.Pipes.PipeDirection]::InOut)
    $pipeClient.Connect(2000)
    $reader = New-Object System.IO.StreamReader($pipeClient)
    $info = $reader.ReadLine() | ConvertFrom-Json
    $slimHubPort = $info.port
    $iframeWsUrl = $info.iframeWsUrl
    $pipeClient.Close()
    Write-Host "SlimHub port: $slimHubPort"
    Write-Host "Iframe WS: $iframeWsUrl"
}

if (-not $iframeWsUrl) {
    Write-Host "Could not get iframe WebSocket URL from MosaicTools" -ForegroundColor Red
    exit 1
}

# Connect to iframe
$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ws.ConnectAsync([Uri]$iframeWsUrl, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
Write-Host "Connected to iframe" -ForegroundColor Green

function Send-CDP($ws, $method, $params = @{}) {
    $id = Get-Random -Minimum 1000 -Maximum 9999
    $msg = @{ id = $id; method = $method; params = $params } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($msg)
    $ws.SendAsync([ArraySegment[byte]]$bytes, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).GetAwaiter().GetResult()

    $buf = New-Object byte[] 65536
    $result = ""
    do {
        $seg = New-Object ArraySegment[byte] $buf
        $recv = $ws.ReceiveAsync($seg, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
        $result += [Text.Encoding]::UTF8.GetString($buf, 0, $recv.Count)
    } while (-not $recv.EndOfMessage)

    return $result | ConvertFrom-Json
}

# Enable Runtime
Send-CDP $ws "Runtime.enable" | Out-Null

# Test 1: Find text "IMPRESSION" and highlight it red
$js = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    if (editors.length < 2 || !editors[1].editor) return 'no editor';
    const editor = editors[1].editor;
    const doc = editor.state.doc;

    // Find the word "unremarkable" anywhere in the doc
    let found = [];
    doc.descendants((node, pos) => {
        if (node.isText) {
            let idx = node.text.toLowerCase().indexOf('unremarkable');
            while (idx !== -1) {
                found.push({ from: pos + idx, to: pos + idx + 'unremarkable'.length });
                idx = node.text.toLowerCase().indexOf('unremarkable', idx + 1);
            }
        }
    });

    if (found.length === 0) return 'word not found';

    // Highlight first occurrence in orange
    const { from, to } = found[0];
    editor.chain()
        .setTextSelection({ from, to })
        .setHighlight({ color: '#FF6B35' })
        .run();

    // Move cursor to end so selection doesn't stay visible
    editor.commands.setTextSelection(doc.content.size - 1);

    return 'highlighted ' + found.length + ' occurrences, first at pos ' + found[0].from;
})()
'@

$result = Send-CDP $ws "Runtime.evaluate" @{ expression = $js; returnByValue = $true }
Write-Host "`nHighlight test result:" -ForegroundColor Cyan
Write-Host ($result.result.result.value)

# Test 2: Remove the highlight we just added
Read-Host "`nPress Enter to UNDO the highlight (restore original)"

$jsUndo = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    if (editors.length < 2 || !editors[1].editor) return 'no editor';
    const editor = editors[1].editor;
    const doc = editor.state.doc;

    let found = [];
    doc.descendants((node, pos) => {
        if (node.isText) {
            let idx = node.text.toLowerCase().indexOf('unremarkable');
            while (idx !== -1) {
                found.push({ from: pos + idx, to: pos + idx + 'unremarkable'.length });
                idx = node.text.toLowerCase().indexOf('unremarkable', idx + 1);
            }
        }
    });

    if (found.length === 0) return 'word not found';

    const { from, to } = found[0];
    editor.chain()
        .setTextSelection({ from, to })
        .unsetHighlight()
        .run();

    editor.commands.setTextSelection(doc.content.size - 1);

    return 'removed highlight from first occurrence';
})()
'@

$result2 = Send-CDP $ws "Runtime.evaluate" @{ expression = $jsUndo; returnByValue = $true }
Write-Host "Undo result:" -ForegroundColor Cyan
Write-Host ($result2.result.result.value)

$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "", [Threading.CancellationToken]::None).GetAwaiter().GetResult()
Write-Host "`nDone." -ForegroundColor Green

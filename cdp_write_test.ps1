$ErrorActionPreference = 'Stop'

$targets = Invoke-RestMethod -Uri 'http://localhost:9222/json'
$iframe = $targets | Where-Object { $_.url -like 'https://rp.radpair.com/*' } | Select-Object -First 1
if (-not $iframe) { Write-Output "No report iframe found"; exit 1 }

$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ct = [System.Threading.CancellationToken]::None
$ws.ConnectAsync([Uri]$iframe.webSocketDebuggerUrl, $ct).Wait()

function Invoke-CDP($id, $js) {
    $payload = @{
        id = $id
        method = 'Runtime.evaluate'
        params = @{
            expression = $js
            returnByValue = $true
        }
    } | ConvertTo-Json -Depth 5 -Compress

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $segment = New-Object System.ArraySegment[byte](,$bytes)
    $ws.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $ct).Wait()

    $buf = New-Object byte[] 262144
    $result = $ws.ReceiveAsync((New-Object System.ArraySegment[byte](,$buf)), $ct).Result
    return [System.Text.Encoding]::UTF8.GetString($buf, 0, $result.Count)
}

# Insert "test " at position 1 (beginning of text) in both editors
$js = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    const results = [];

    editors.forEach((pm, idx) => {
        if (!pm.editor) {
            results.push({ index: idx, error: 'no editor' });
            return;
        }

        const editor = pm.editor;
        const before = editor.getText().substring(0, 80);

        // Insert "test " at position 1 (start of document content, after doc node opening)
        editor.commands.insertContentAt(1, 'test ');

        const after = editor.getText().substring(0, 80);
        results.push({
            index: idx,
            before: before,
            after: after,
            success: true
        });
    });

    return JSON.stringify(results, null, 2);
})()
'@

Write-Output "=== Insert 'test ' at beginning of both editors ==="
$r = Invoke-CDP 1 $js
$parsed = ($r | ConvertFrom-Json)
if ($parsed.result.result.value) {
    Write-Output $parsed.result.result.value
} else {
    Write-Output $r
}

$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', $ct).Wait()

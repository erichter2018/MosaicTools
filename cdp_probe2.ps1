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

# Probe: Explore the TipTap editor object
$js = @'
(() => {
    const pm = document.querySelector('.ProseMirror');
    if (!pm || !pm.editor) return JSON.stringify({ error: 'No editor found' });

    const editor = pm.editor;
    const result = {};

    // Basic editor info
    result.isEditable = editor.isEditable;
    result.isFocused = editor.isFocused;
    result.isEmpty = editor.isEmpty;

    // Available commands (TipTap API)
    result.commandNames = Object.keys(editor.commands || {}).sort().slice(0, 50);

    // Get document as text
    result.plainText = editor.getText();

    // Get document as JSON structure (first 2 levels)
    const json = editor.getJSON();
    result.docType = json.type;
    result.topLevelNodes = (json.content || []).map(n => ({
        type: n.type,
        textPreview: (n.content || []).map(c => c.text || c.type).join('').substring(0, 80)
    }));

    // Editor state info
    const state = editor.state;
    result.docSize = state.doc.content.size;
    result.selectionFrom = state.selection.from;
    result.selectionTo = state.selection.to;
    result.selectionEmpty = state.selection.empty;

    // Available extensions
    result.extensions = editor.extensionManager.extensions.map(e => e.name).sort();

    // Can we insert content?
    result.hasInsertContent = typeof editor.commands.insertContent === 'function';
    result.hasInsertContentAt = typeof editor.commands.insertContentAt === 'function';
    result.hasSetContent = typeof editor.commands.setContent === 'function';

    return JSON.stringify(result, null, 2);
})()
'@

Write-Output "=== TipTap Editor Exploration ==="
$r = Invoke-CDP 1 $js
$parsed = ($r | ConvertFrom-Json)
if ($parsed.result.result.value) {
    Write-Output $parsed.result.result.value
} else {
    Write-Output $r
}

$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', $ct).Wait()

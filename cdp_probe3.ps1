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

# Find ALL ProseMirror editors on the page
$js = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    const results = [];

    editors.forEach((pm, idx) => {
        const info = {
            index: idx,
            hasEditor: !!pm.editor,
            classList: Array.from(pm.classList),
            contentEditable: pm.contentEditable,
            textPreview: pm.innerText.substring(0, 200),
            textLength: pm.innerText.length,
            isFocused: pm.classList.contains('ProseMirror-focused')
        };

        // Walk up to find context (what section is this editor in?)
        let el = pm.parentElement;
        let depth = 0;
        let contextHints = [];
        while (el && depth < 15) {
            const cls = (el.className || '').toString();
            // Look for section identifiers
            if (cls.includes('transcript') || cls.includes('report') || cls.includes('finding')
                || cls.includes('impression') || cls.includes('editor') || cls.includes('section')
                || cls.includes('RichText') || cls.includes('Tiptap')) {
                contextHints.push({
                    depth: depth,
                    tag: el.tagName,
                    className: cls.substring(0, 120),
                    id: el.id || null,
                    ariaLabel: el.getAttribute('aria-label') || null,
                    role: el.getAttribute('role') || null
                });
            }
            depth++;
            el = el.parentElement;
        }
        info.contextHints = contextHints;

        // Check for labels/headers near this editor
        const parent = pm.closest('[class*="MuiTiptap"]') || pm.parentElement;
        if (parent) {
            const prevSibling = parent.previousElementSibling;
            if (prevSibling) {
                info.prevSiblingText = prevSibling.innerText ? prevSibling.innerText.substring(0, 100) : null;
                info.prevSiblingClass = (prevSibling.className || '').substring(0, 100);
            }
        }

        if (pm.editor) {
            const state = pm.editor.state;
            info.docSize = state.doc.content.size;
            info.isEditable = pm.editor.isEditable;
            info.extensions = pm.editor.extensionManager.extensions.map(e => e.name).sort();

            // Get top-level node types
            const json = pm.editor.getJSON();
            info.topNodes = (json.content || []).slice(0, 5).map(n => ({
                type: n.type,
                text: (n.content || []).map(c => c.text || '').join('').substring(0, 60)
            }));
        }

        results.push(info);
    });

    return JSON.stringify(results, null, 2);
})()
'@

Write-Output "=== All ProseMirror Editors ==="
$r = Invoke-CDP 1 $js
$parsed = ($r | ConvertFrom-Json)
if ($parsed.result.result.value) {
    Write-Output $parsed.result.result.value
} else {
    Write-Output $r
}

$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', $ct).Wait()

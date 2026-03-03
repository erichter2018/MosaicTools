$ErrorActionPreference = 'Stop'

# Get the report iframe target
$targets = Invoke-RestMethod -Uri 'http://localhost:9222/json'
$iframe = $targets | Where-Object { $_.url -like 'https://rp.radpair.com/*' } | Select-Object -First 1
if (-not $iframe) {
    Write-Output "No report iframe found"
    exit 1
}
Write-Output "Target: $($iframe.url)"
$wsUrl = $iframe.webSocketDebuggerUrl

# Connect WebSocket
$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ct = [System.Threading.CancellationToken]::None
$ws.ConnectAsync([Uri]$wsUrl, $ct).Wait()

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

    $buf = New-Object byte[] 131072
    $result = $ws.ReceiveAsync((New-Object System.ArraySegment[byte](,$buf)), $ct).Result
    $resp = [System.Text.Encoding]::UTF8.GetString($buf, 0, $result.Count)
    return $resp
}

# Probe 1: Find ProseMirror and check for view reference
$js1 = @'
(() => {
    const pm = document.querySelector('.ProseMirror');
    if (!pm) return JSON.stringify({ error: 'No ProseMirror element found' });

    // Check various ways ProseMirror view might be accessible
    const checks = {
        hasPmViewDesc: !!pm.pmViewDesc,
        hasEditorView: !!window.editorView,
        hasView: !!pm.__view,
        classList: Array.from(pm.classList),
        contentEditable: pm.contentEditable,
        textLength: pm.innerText.length,
        textPreview: pm.innerText.substring(0, 300)
    };

    // Try to find EditorView by walking up from ProseMirror element
    let el = pm;
    let viewKeys = [];
    while (el) {
        const keys = Object.keys(el).filter(k => k.toLowerCase().includes('view') || k.toLowerCase().includes('editor') || k.toLowerCase().includes('prose'));
        if (keys.length > 0) viewKeys.push({ tag: el.tagName, id: el.id, keys: keys });
        el = el.parentElement;
    }
    checks.viewKeysInAncestors = viewKeys;

    // Check window-level globals
    const globals = Object.keys(window).filter(k =>
        k.toLowerCase().includes('editor') ||
        k.toLowerCase().includes('prose') ||
        k.toLowerCase().includes('view') ||
        k.toLowerCase().includes('tiptap')
    );
    checks.editorGlobals = globals.slice(0, 20);

    return JSON.stringify(checks, null, 2);
})()
'@

Write-Output "`n=== Probe 1: ProseMirror Detection ==="
$r1 = Invoke-CDP 1 $js1
$parsed = $r1 | ConvertFrom-Json
if ($parsed.result.result.value) {
    Write-Output $parsed.result.result.value
} else {
    Write-Output $r1
}

# Probe 2: Try to find the ProseMirror EditorView through internal properties
$js2 = @'
(() => {
    const pm = document.querySelector('.ProseMirror');
    if (!pm) return 'no pm';

    // ProseMirror stores the view in a non-enumerable property
    // Check all own properties including non-enumerable
    const allProps = Object.getOwnPropertyNames(pm);
    const interestingProps = allProps.filter(p => !p.startsWith('__') || p.includes('view') || p.includes('editor'));

    // Also check for Symbol properties
    const symbols = Object.getOwnPropertySymbols(pm);

    // Try the known ProseMirror internal: pm.cmView or similar
    // TipTap/ProseMirror often stores view reference
    let viewFound = null;
    for (const prop of allProps) {
        try {
            const val = pm[prop];
            if (val && typeof val === 'object' && val.state && val.dispatch) {
                viewFound = prop;
                break;
            }
        } catch(e) {}
    }

    // Check if it's a TipTap editor (wrapper around ProseMirror)
    const allWindowProps = Object.getOwnPropertyNames(window);
    const storeProps = allWindowProps.filter(p => p.includes('store') || p.includes('app') || p.includes('__'));

    return JSON.stringify({
        ownPropCount: allProps.length,
        interestingProps: interestingProps.slice(0, 30),
        symbolCount: symbols.length,
        viewFoundAt: viewFound,
        windowStoreHints: storeProps.slice(0, 20)
    }, null, 2);
})()
'@

Write-Output "`n=== Probe 2: EditorView Search ==="
$r2 = Invoke-CDP 2 $js2
$parsed2 = $r2 | ConvertFrom-Json
if ($parsed2.result.result.value) {
    Write-Output $parsed2.result.result.value
} else {
    Write-Output $r2
}

# Probe 3: Check React/framework internals (Mosaic might use React)
$js3 = @'
(() => {
    const pm = document.querySelector('.ProseMirror');
    if (!pm) return 'no pm';

    // React stores internal fiber on __reactFiber or __reactInternalInstance
    const reactKeys = Object.keys(pm).filter(k => k.startsWith('__react'));

    // Check parent elements for React/Angular/Vue bindings
    let el = pm.parentElement;
    let depth = 0;
    let frameworkHints = [];
    while (el && depth < 10) {
        const rKeys = Object.keys(el).filter(k => k.startsWith('__react') || k.startsWith('__ng') || k.startsWith('__vue'));
        if (rKeys.length) frameworkHints.push({ depth, tag: el.tagName, className: (el.className||'').substring(0,80), keys: rKeys });
        depth++;
        el = el.parentElement;
    }

    // Check for ProseMirror view via DOM event listeners
    // The view attaches event listeners we might trace
    const hasInputListener = typeof pm.oninput === 'function';

    return JSON.stringify({
        pmReactKeys: reactKeys,
        frameworkHints,
        hasOninput: hasInputListener,
        pmParentTag: pm.parentElement ? pm.parentElement.tagName : null,
        pmParentClass: pm.parentElement ? (pm.parentElement.className||'').substring(0,100) : null
    }, null, 2);
})()
'@

Write-Output "`n=== Probe 3: Framework Detection ==="
$r3 = Invoke-CDP 3 $js3
$parsed3 = $r3 | ConvertFrom-Json
if ($parsed3.result.result.value) {
    Write-Output $parsed3.result.result.value
} else {
    Write-Output $r3
}

$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', $ct).Wait()
Write-Output "`nDone."

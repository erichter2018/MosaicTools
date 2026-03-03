$ErrorActionPreference = 'Stop'

$port = (Get-Content 'C:\Users\erik.richter\AppData\Local\Packages\MosaicInfoHub_yzjagrpanxm4c\LocalState\EBWebView\DevToolsActivePort' | Select-Object -First 1).Trim()
Write-Output "CDP Port: $port"

$targets = Invoke-RestMethod -Uri "http://localhost:$port/json"

function Invoke-CDPTarget($wsUrl, $id, $js) {
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $ct = [System.Threading.CancellationToken]::None
    $ws.ConnectAsync([Uri]$wsUrl, $ct).Wait()

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

    $buf = New-Object byte[] 524288
    $result = $ws.ReceiveAsync((New-Object System.ArraySegment[byte](,$buf)), $ct).Result
    $resp = [System.Text.Encoding]::UTF8.GetString($buf, 0, $result.Count)
    $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', $ct).Wait()
    return $resp
}

# === Probe the MAIN SlimHub page ===
$main = $targets | Where-Object { $_.title -eq 'SlimHub' } | Select-Object -First 1

$jsMain = @'
(() => {
    const result = {};

    // 1. Look for patient info, accession, study description in the DOM
    // Search for text content that looks like metadata
    const allText = document.body.innerText;

    // 2. Check for React state / Next.js data
    if (window.__NEXT_DATA__) {
        result.nextData = JSON.stringify(window.__NEXT_DATA__).substring(0, 500);
    }

    // 3. Look for specific UI elements by class/role
    const labels = document.querySelectorAll('[class*="patient"], [class*="accession"], [class*="study"], [class*="mrn"], [class*="description"], [aria-label]');
    result.labeledElements = [];
    labels.forEach(el => {
        result.labeledElements.push({
            tag: el.tagName,
            class: (el.className || '').toString().substring(0, 100),
            ariaLabel: el.getAttribute('aria-label'),
            text: (el.innerText || '').substring(0, 150),
            id: el.id || null
        });
    });

    // 4. Check for any global state stores (Redux, etc.)
    const storeKeys = Object.keys(window).filter(k =>
        k.includes('store') || k.includes('Store') || k.includes('redux') || k.includes('state') || k.includes('State')
    );
    result.storeKeys = storeKeys;

    // 5. Grab visible text that might contain metadata (first 2000 chars)
    result.visibleTextPreview = allText.substring(0, 2000);

    return JSON.stringify(result, null, 2);
})()
'@

Write-Output "`n=== SlimHub Main Page ==="
$r1 = Invoke-CDPTarget $main.webSocketDebuggerUrl 1 $jsMain
$parsed1 = ($r1 | ConvertFrom-Json)
if ($parsed1.result.result.value) {
    Write-Output $parsed1.result.result.value
} else {
    Write-Output $r1
}

# === Probe the report iframe ===
$iframe = $targets | Where-Object { $_.url -like 'https://rp.radpair.com/*' } | Select-Object -First 1

$jsIframe = @'
(() => {
    const result = {};

    // 1. Look for metadata elements
    const allElements = document.querySelectorAll('[class*="patient"], [class*="accession"], [class*="study"], [class*="mrn"], [class*="description"], [class*="template"], [class*="detail"], [class*="header"], [class*="status"]');
    result.metadataElements = [];
    allElements.forEach(el => {
        const text = (el.innerText || '').substring(0, 200);
        if (text.length > 0) {
            result.metadataElements.push({
                tag: el.tagName,
                class: (el.className || '').toString().substring(0, 120),
                text: text,
                id: el.id || null
            });
        }
    });

    // 2. Check Next.js / React data stores
    if (window.__NEXT_DATA__) {
        const nd = window.__NEXT_DATA__;
        result.nextDataProps = nd.props ? Object.keys(nd.props) : [];
        result.nextDataPage = nd.page;
        result.nextDataQuery = nd.query;
        // Try to get page props (often has the actual data)
        if (nd.props && nd.props.pageProps) {
            result.pageProps = JSON.stringify(nd.props.pageProps).substring(0, 2000);
        }
    }

    // 3. Check for global state
    const globals = Object.keys(window).filter(k =>
        k.includes('store') || k.includes('Store') || k.includes('redux') || k.includes('state')
        || k.includes('__APOLLO') || k.includes('__RELAY') || k.includes('__TANSTACK')
    );
    result.stateGlobals = globals;

    // 4. Look for radpair-specific classes
    const rpElements = document.querySelectorAll('[class*="radpair"]');
    result.radpairElements = [];
    rpElements.forEach(el => {
        const text = (el.innerText || '').substring(0, 200);
        result.radpairElements.push({
            tag: el.tagName,
            class: (el.className || '').toString().substring(0, 120),
            textPreview: text.substring(0, 100),
            childCount: el.children.length
        });
    });

    // 5. Look for the Study/template selector
    const selects = document.querySelectorAll('select, [role="combobox"], [role="listbox"], [class*="Select"], [class*="select"], [class*="Study"], [class*="study"]');
    result.selectElements = [];
    selects.forEach(el => {
        result.selectElements.push({
            tag: el.tagName,
            class: (el.className || '').toString().substring(0, 120),
            role: el.getAttribute('role'),
            ariaLabel: el.getAttribute('aria-label'),
            value: el.value || (el.innerText || '').substring(0, 100)
        });
    });

    // 6. Page URL and title
    result.url = window.location.href;
    result.title = document.title;

    // 7. Visible text (first 3000 chars)
    result.visibleText = document.body.innerText.substring(0, 3000);

    return JSON.stringify(result, null, 2);
})()
'@

Write-Output "`n=== Report Iframe (rp.radpair.com) ==="
$r2 = Invoke-CDPTarget $iframe.webSocketDebuggerUrl 1 $jsIframe
$parsed2 = ($r2 | ConvertFrom-Json)
if ($parsed2.result.result.value) {
    Write-Output $parsed2.result.result.value
} else {
    Write-Output $r2
}

Write-Output "`nDone."

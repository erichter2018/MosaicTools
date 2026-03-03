$ErrorActionPreference = 'Stop'

# Find CDP port from MosaicInfoHub DevToolsActivePort file (same as CdpService.cs)
$packagesDir = Join-Path $env:LOCALAPPDATA 'Packages'
$mosaicDir = Get-ChildItem $packagesDir -Directory -Filter 'MosaicInfoHub_*' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $mosaicDir) { Write-Output "MosaicInfoHub package not found"; exit 1 }
$portFile = Join-Path $mosaicDir.FullName 'LocalState\EBWebView\DevToolsActivePort'
if (-not (Test-Path $portFile)) { Write-Output "DevToolsActivePort file not found at $portFile"; exit 1 }
$port = (Get-Content $portFile)[0].Trim()
Write-Output "CDP port: $port (from $portFile)"

$targets = Invoke-RestMethod -Uri "http://localhost:$port/json"
Write-Output "`n=== ALL CDP TARGETS ==="
foreach ($t in $targets) {
    Write-Output "  [$($t.type)] $($t.title) => $($t.url)"
}

$iframe = $targets | Where-Object { $_.url -like 'https://rp.radpair.com/*' } | Select-Object -First 1
if (-not $iframe) { Write-Output "No report iframe found"; exit 1 }

$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ct = [System.Threading.CancellationToken]::None
$ws.ConnectAsync([Uri]$iframe.webSocketDebuggerUrl, $ct).Wait()

$global:cmdId = 0

function Invoke-CDP($js) {
    $global:cmdId++
    $payload = @{
        id = $global:cmdId
        method = 'Runtime.evaluate'
        params = @{
            expression = $js
            returnByValue = $true
            awaitPromise = $false
        }
    } | ConvertTo-Json -Depth 5 -Compress

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $segment = New-Object System.ArraySegment[byte](,$bytes)
    $ws.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $ct).Wait()

    # Read response — handle multi-frame messages
    $allBytes = New-Object System.Collections.Generic.List[byte]
    do {
        $buf = New-Object byte[] 524288
        $result = $ws.ReceiveAsync((New-Object System.ArraySegment[byte](,$buf)), $ct).Result
        for ($i = 0; $i -lt $result.Count; $i++) { $allBytes.Add($buf[$i]) }
    } while (-not $result.EndOfMessage)

    $text = [System.Text.Encoding]::UTF8.GetString($allBytes.ToArray())
    $parsed = $text | ConvertFrom-Json
    if ($parsed.result.result.value) {
        return $parsed.result.result.value
    }
    if ($parsed.result.result.description) {
        return "ERROR: $($parsed.result.result.description)"
    }
    return $text
}

$output = [System.Text.StringBuilder]::new()
function Log($text) {
    $output.AppendLine($text) | Out-Null
    Write-Output $text
}

Log "# Mosaic CDP Full Inventory"
Log "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log ""

# ============================================================
Log "## 1. PROSEMIRROR EDITORS & EXTENSIONS"
Log ""
# ============================================================

$js = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    const results = [];
    editors.forEach((pm, idx) => {
        const info = { index: idx, hasEditor: !!pm.editor, editable: pm.contentEditable };
        if (pm.editor) {
            info.isEditable = pm.editor.isEditable;
            info.docSize = pm.editor.state.doc.content.size;
            info.extensions = pm.editor.extensionManager.extensions.map(e => ({
                name: e.name,
                type: e.type,
                // Check for options/config
                hasOptions: e.options ? Object.keys(e.options).length : 0
            }));
            // Check for custom commands beyond standard
            info.commands = Object.keys(pm.editor.commands).sort();
            // Check for marks (bold, italic, etc.)
            const schema = pm.editor.state.schema;
            info.marks = Object.keys(schema.marks);
            info.nodeTypes = Object.keys(schema.nodes);
            // Storage - extensions store data here
            info.storageKeys = Object.keys(pm.editor.extensionStorage || {});
            // Check for input rules (auto-text triggers)
            const plugins = pm.editor.state.plugins;
            info.pluginCount = plugins.length;
            info.pluginKeys = plugins.map(p => p.key).filter(k => k).slice(0, 30);
        }
        // Parent context
        let el = pm.parentElement;
        let labels = [];
        for (let d = 0; d < 10 && el; d++) {
            const lbl = el.getAttribute('aria-label') || el.getAttribute('data-testid') || '';
            if (lbl) labels.push(lbl);
            el = el.parentElement;
        }
        info.parentLabels = labels;
        results.push(info);
    });
    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 2. WINDOW/GLOBAL OBJECT EXPLORATION"
Log ""
# ============================================================

$js = @'
(() => {
    // Look for app state, stores, APIs on window
    const dominated = new Set([
        'window','self','document','location','navigator','performance','screen',
        'history','localStorage','sessionStorage','crypto','indexedDB','caches',
        'console','fetch','alert','confirm','prompt','open','close','focus','blur',
        'setTimeout','setInterval','clearTimeout','clearInterval','requestAnimationFrame',
        'cancelAnimationFrame','getComputedStyle','matchMedia','postMessage',
        'addEventListener','removeEventListener','dispatchEvent','atob','btoa',
        'URL','URLSearchParams','Blob','File','FileReader','FormData',
        'XMLHttpRequest','WebSocket','Worker','SharedWorker','ServiceWorker',
        'MutationObserver','IntersectionObserver','ResizeObserver',
        'Promise','Map','Set','WeakMap','WeakSet','Proxy','Reflect',
        'JSON','Math','Date','RegExp','Error','Array','Object','String',
        'Number','Boolean','Symbol','Function','Infinity','NaN','undefined',
        'parseInt','parseFloat','isNaN','isFinite','eval','encodeURI','decodeURI',
        'encodeURIComponent','decodeURIComponent','escape','unescape',
        'globalThis','structuredClone','queueMicrotask','reportError',
        'length','name','customElements','origin','closed','frames','parent','top',
        'frameElement','opener','status','toolbar','menubar','personalbar',
        'scrollbars','statusbar','locationbar','visualViewport','navigation',
        'devicePixelRatio','innerWidth','innerHeight','outerWidth','outerHeight',
        'screenX','screenY','screenLeft','screenTop','pageXOffset','pageYOffset',
        'scrollX','scrollY','clientInformation','styleMedia','onload','onerror',
        'onmessage','onbeforeunload','onunload','onhashchange','onpopstate',
        'onfocus','onblur','onresize','onscroll','onmousedown','onmouseup',
        'onmousemove','onmouseover','onmouseout','onclick','ondblclick',
        'onkeydown','onkeyup','onkeypress','ontouchstart','ontouchend',
        'ontouchmove','ontouchcancel','oncontextmenu','onwheel','onpointerdown',
        'onpointerup','onpointermove','onpointerover','onpointerout',
        'onpointerenter','onpointerleave','onpointercancel','ongotpointercapture',
        'onlostpointercapture','onanimationend','onanimationiteration',
        'onanimationstart','ontransitionend','ontransitionrun','ontransitionstart',
        'ontransitioncancel','onabort','oncanplay','oncanplaythrough',
        'ondurationchange','onemptied','onended','oninput','oninvalid',
        'onloadeddata','onloadedmetadata','onloadstart','onpause','onplay',
        'onplaying','onprogress','onratechange','onreset','onseeked','onseeking',
        'onselect','onstalled','onsubmit','onsuspend','ontimeupdate',
        'onvolumechange','onwaiting','onchange','onstorage','onbeforeinput',
        'onformdata','onsecuritypolicyviolation','onslotchange','oncuechange',
        'speechSynthesis','getSelection','find','print','stop','moveBy','moveTo',
        'resizeBy','resizeTo','scroll','scrollTo','scrollBy','requestIdleCallback',
        'cancelIdleCallback','showDirectoryPicker','showOpenFilePicker',
        'showSaveFilePicker','getScreenDetails','queryLocalFonts',
        'createImageBitmap','chrome','webkitRequestAnimationFrame',
        'webkitCancelAnimationFrame','getDigitalGoodsService',
        'launchQueue','documentPictureInPicture','onpageswap','onpagereveal',
        'onpageshow','onpagehide','onbeforematch','onbeforetoggle','ontoggle',
        'crossOriginIsolated','scheduler','trustedTypes','cookieStore',
        'oncontentvisibilityautostatechange','onscrollend','onscrollsnapchanging',
        'onscrollsnapchange','ai','model','isSecureContext','cdc_adoQpoasnfa76pfcZLmcfl_Array',
        'cdc_adoQpoasnfa76pfcZLmcfl_JSON','cdc_adoQpoasnfa76pfcZLmcfl_Object',
        'cdc_adoQpoasnfa76pfcZLmcfl_Promise','cdc_adoQpoasnfa76pfcZLmcfl_Proxy',
        'cdc_adoQpoasnfa76pfcZLmcfl_Symbol'
    ]);
    const interesting = {};
    for (const key of Object.getOwnPropertyNames(window)) {
        if (dominated.has(key)) continue;
        if (key.startsWith('on')) continue;  // event handlers
        if (key.startsWith('webkit') || key.startsWith('__coverage')) continue;
        try {
            const val = window[key];
            const type = typeof val;
            if (type === 'function') {
                interesting[key] = { type: 'function', length: val.length };
            } else if (type === 'object' && val !== null) {
                const keys = Object.keys(val).slice(0, 20);
                interesting[key] = { type: 'object', keys: keys, proto: val.constructor?.name || '?' };
            } else if (type === 'string' || type === 'number' || type === 'boolean') {
                interesting[key] = { type, value: String(val).substring(0, 200) };
            }
        } catch(e) {
            interesting[key] = { type: 'error', msg: e.message };
        }
    }
    return JSON.stringify(interesting, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 3. REACT INTERNALS"
Log ""
# ============================================================

$js = @'
(() => {
    const results = {};
    // Check for React DevTools hook
    results.hasReactDevTools = !!window.__REACT_DEVTOOLS_GLOBAL_HOOK__;
    // Check for Next.js data
    results.hasNextData = !!window.__NEXT_DATA__;
    if (window.__NEXT_DATA__) {
        results.nextData = {
            buildId: window.__NEXT_DATA__.buildId,
            page: window.__NEXT_DATA__.page,
            propsKeys: Object.keys(window.__NEXT_DATA__.props || {}),
            runtimeConfig: window.__NEXT_DATA__.runtimeConfig
        };
    }
    // Check for Redux/MobX/Zustand stores
    results.hasReduxStore = !!window.__REDUX_DEVTOOLS_EXTENSION__;
    results.hasReduxStore2 = !!window.store;
    // Try to find React root
    const rootEl = document.getElementById('root') || document.getElementById('app') || document.getElementById('__next');
    results.rootElement = rootEl ? { id: rootEl.id, tag: rootEl.tagName } : null;
    // Look for React fiber on root
    if (rootEl) {
        const fiberKey = Object.keys(rootEl).find(k => k.startsWith('__reactFiber') || k.startsWith('__reactInternalInstance'));
        results.hasFiber = !!fiberKey;
        results.fiberKey = fiberKey || null;
        if (fiberKey) {
            // Walk fiber tree to find state stores
            const fiber = rootEl[fiberKey];
            let node = fiber;
            const stateNodes = [];
            let depth = 0;
            while (node && depth < 50) {
                if (node.memoizedState && node.type) {
                    const name = node.type.displayName || node.type.name || '?';
                    if (name !== '?' && name.length > 1) {
                        stateNodes.push(name);
                    }
                }
                node = node.child || node.sibling || node.return;
                depth++;
            }
            results.fiberComponents = [...new Set(stateNodes)].slice(0, 30);
        }
    }
    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 4. ALL BUTTONS & INTERACTIVE CONTROLS"
Log ""
# ============================================================

$js = @'
(() => {
    const controls = [];
    // All buttons
    document.querySelectorAll('button').forEach(b => {
        controls.push({
            type: 'button',
            text: (b.innerText || '').trim().substring(0, 80),
            ariaLabel: b.getAttribute('aria-label') || null,
            disabled: b.disabled,
            className: (b.className || '').substring(0, 100),
            id: b.id || null,
            dataTestId: b.getAttribute('data-testid') || null,
            visible: b.offsetParent !== null
        });
    });
    // All inputs
    document.querySelectorAll('input, select, textarea').forEach(inp => {
        controls.push({
            type: inp.tagName.toLowerCase() + ':' + (inp.type || ''),
            id: inp.id || null,
            name: inp.name || null,
            ariaLabel: inp.getAttribute('aria-label') || null,
            placeholder: inp.placeholder || null,
            role: inp.getAttribute('role') || null,
            value: (inp.value || '').substring(0, 100),
            disabled: inp.disabled,
            dataTestId: inp.getAttribute('data-testid') || null
        });
    });
    // Comboboxes / listboxes
    document.querySelectorAll('[role="combobox"], [role="listbox"], [role="menu"], [role="menubar"], [role="tablist"]').forEach(el => {
        controls.push({
            type: 'aria:' + el.getAttribute('role'),
            tag: el.tagName,
            ariaLabel: el.getAttribute('aria-label') || null,
            text: (el.innerText || '').substring(0, 100),
            id: el.id || null
        });
    });
    return JSON.stringify(controls, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 5. KEYBOARD EVENT HANDLERS & INPUT RULES"
Log ""
# ============================================================

$js = @'
(() => {
    const results = {};
    // Check ProseMirror plugins for input rules (autotext triggers)
    const editors = document.querySelectorAll('.ProseMirror');
    editors.forEach((pm, idx) => {
        if (!pm.editor) return;
        const editorInfo = { plugins: [] };
        pm.editor.state.plugins.forEach(plugin => {
            const info = { key: plugin.key || '(unnamed)' };
            // Check for inputRules
            if (plugin.spec && plugin.spec.rules) {
                info.inputRules = plugin.spec.rules.map(r => ({
                    regex: r.match ? r.match.toString() : null,
                    hasHandler: !!r.handler
                }));
            }
            // Check for keymap/handleKeyDown
            if (plugin.props) {
                info.hasHandleKeyDown = !!plugin.props.handleKeyDown;
                info.hasHandlePaste = !!plugin.props.handlePaste;
                info.hasHandleTextInput = !!plugin.props.handleTextInput;
                info.hasDecorations = !!plugin.props.decorations;
                info.propKeys = Object.keys(plugin.props);
            }
            if (info.inputRules || info.hasHandleKeyDown || info.hasHandlePaste || info.hasHandleTextInput) {
                editorInfo.plugins.push(info);
            }
        });
        results['editor_' + idx] = editorInfo;
    });
    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 6. PROSEMIRROR COMMANDS DETAIL"
Log ""
# ============================================================

$js = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    if (editors.length === 0 || !editors[0].editor) return 'no editors';
    const editor = editors[0].editor;
    // All available commands
    const commands = Object.keys(editor.commands).sort();
    // Group by likely category
    const result = {
        allCommands: commands,
        totalCount: commands.length,
        // Check for macro/autotext related
        macroRelated: commands.filter(c => /macro|auto|text|template|insert|expand|snippet|trigger/i.test(c)),
        // Selection related
        selection: commands.filter(c => /select|cursor|focus|blur/i.test(c)),
        // Content manipulation
        content: commands.filter(c => /insert|delete|replace|set|clear|toggle|update|split|merge|join|lift|sink|wrap/i.test(c)),
        // History
        history: commands.filter(c => /undo|redo|history/i.test(c))
    };
    return JSON.stringify(result, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 7. EXTENSION STORAGE & STATE"
Log ""
# ============================================================

$js = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    const results = {};
    editors.forEach((pm, idx) => {
        if (!pm.editor) return;
        const storage = pm.editor.extensionStorage;
        const info = {};
        for (const [key, val] of Object.entries(storage)) {
            try {
                info[key] = {
                    keys: Object.keys(val || {}),
                    preview: JSON.stringify(val).substring(0, 300)
                };
            } catch(e) {
                info[key] = { error: e.message };
            }
        }
        results['editor_' + idx] = info;
    });
    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 8. DOM TREE STRUCTURE (HIGH LEVEL)"
Log ""
# ============================================================

$js = @'
(() => {
    // Walk the DOM at a high level to understand page structure
    function walk(el, depth) {
        if (depth > 6) return null;
        const info = {
            tag: el.tagName,
            id: el.id || undefined,
            role: el.getAttribute('role') || undefined,
            ariaLabel: el.getAttribute('aria-label') || undefined,
            dataTestId: el.getAttribute('data-testid') || undefined,
            className: el.className ? (typeof el.className === 'string' ? el.className : '').substring(0, 80) : undefined
        };
        // Only expand if it has interesting attributes or is a container
        const children = [];
        for (const child of el.children) {
            if (child.tagName === 'SCRIPT' || child.tagName === 'STYLE' || child.tagName === 'LINK') continue;
            const childInfo = walk(child, depth + 1);
            if (childInfo) children.push(childInfo);
        }
        if (children.length > 0) info.children = children.slice(0, 20);
        // Skip boring nodes with no attributes
        if (!info.id && !info.role && !info.ariaLabel && !info.dataTestId
            && (!info.className || info.className.length < 3)
            && children.length === 0 && depth > 2) return null;
        return info;
    }
    return JSON.stringify(walk(document.body, 0), null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 9. LOCALSTORAGE & SESSIONSTORAGE"
Log ""
# ============================================================

$js = @'
(() => {
    const result = { localStorage: {}, sessionStorage: {} };
    try {
        for (let i = 0; i < localStorage.length && i < 50; i++) {
            const key = localStorage.key(i);
            const val = localStorage.getItem(key);
            result.localStorage[key] = val ? val.substring(0, 300) : null;
        }
    } catch(e) { result.localStorageError = e.message; }
    try {
        for (let i = 0; i < sessionStorage.length && i < 50; i++) {
            const key = sessionStorage.key(i);
            const val = sessionStorage.getItem(key);
            result.sessionStorage[key] = val ? val.substring(0, 300) : null;
        }
    } catch(e) { result.sessionStorageError = e.message; }
    return JSON.stringify(result, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 10. EVENT LISTENERS ON KEY ELEMENTS"
Log ""
# ============================================================

$js = @'
(() => {
    // Check what event listeners are registered via getEventListeners (Chrome DevTools API)
    // Note: getEventListeners is only available in Chrome DevTools console, not in CDP Runtime.evaluate
    // Instead, check for jQuery/React event delegation patterns
    const results = {};
    // Check document-level listeners
    const docEvents = [];
    const eventTypes = ['keydown','keyup','keypress','input','paste','copy','cut','beforeinput',
                         'compositionstart','compositionend','textInput'];
    // Check if there are keyboard shortcuts registered
    const editors = document.querySelectorAll('.ProseMirror');
    editors.forEach((pm, idx) => {
        if (!pm.editor) return;
        // Check for keyboard shortcuts in extensions
        const shortcuts = {};
        pm.editor.extensionManager.extensions.forEach(ext => {
            if (ext.options && ext.options.shortcuts) {
                shortcuts[ext.name] = ext.options.shortcuts;
            }
            // Tiptap extensions can define addKeyboardShortcuts()
            if (ext.type === 'extension' || ext.type === 'node' || ext.type === 'mark') {
                try {
                    const parent = ext.parent;
                    // Try to access the keyboard shortcuts config
                    if (ext.config && ext.config.addKeyboardShortcuts) {
                        shortcuts[ext.name + '_config'] = 'has addKeyboardShortcuts';
                    }
                } catch(e) {}
            }
        });
        results['editor_' + idx + '_shortcuts'] = shortcuts;
    });
    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 11. FETCH/XHR API ENDPOINTS (from performance entries)"
Log ""
# ============================================================

$js = @'
(() => {
    // Get all network requests from Performance API
    const entries = performance.getEntriesByType('resource');
    const apiCalls = entries.filter(e =>
        e.initiatorType === 'fetch' || e.initiatorType === 'xmlhttprequest'
    ).map(e => ({
        url: e.name.substring(0, 200),
        type: e.initiatorType,
        duration: Math.round(e.duration),
        size: e.transferSize || 0
    }));
    // Deduplicate by URL path (ignore query params)
    const seen = new Set();
    const unique = [];
    apiCalls.forEach(c => {
        try {
            const u = new URL(c.url);
            const key = u.origin + u.pathname;
            if (!seen.has(key)) {
                seen.add(key);
                unique.push({ ...c, path: u.pathname });
            }
        } catch(e) {
            unique.push(c);
        }
    });
    return JSON.stringify(unique, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 12. MACRO/AUTOTEXT DEEP PROBE"
Log ""
# ============================================================

$js = @'
(() => {
    const results = {};

    // 1. Search for macro-related properties on window
    const macroKeys = [];
    for (const key of Object.getOwnPropertyNames(window)) {
        if (/macro|autotext|auto.?text|snippet|template|abbreviat|expand|shortcut|phrase/i.test(key)) {
            macroKeys.push(key);
        }
    }
    results.windowMacroKeys = macroKeys;

    // 2. Search React component tree for macro-related components
    const rootEl = document.getElementById('root') || document.getElementById('app') || document.getElementById('__next');
    if (rootEl) {
        const fiberKey = Object.keys(rootEl).find(k => k.startsWith('__reactFiber'));
        if (fiberKey) {
            const macroComponents = [];
            const visited = new Set();
            function walkFiber(node, depth) {
                if (!node || depth > 100 || visited.has(node)) return;
                visited.add(node);
                if (node.type) {
                    const name = node.type.displayName || node.type.name || '';
                    if (/macro|autotext|auto.?text|snippet|template|abbreviat|expand|shortcut|phrase|dictation|speech|voice|dragon/i.test(name)) {
                        macroComponents.push({
                            name: name,
                            hasState: !!node.memoizedState,
                            propsKeys: node.memoizedProps ? Object.keys(node.memoizedProps).slice(0, 10) : []
                        });
                    }
                }
                // Also check memoizedProps for macro-related data
                if (node.memoizedProps) {
                    for (const [k, v] of Object.entries(node.memoizedProps)) {
                        if (/macro|autotext|snippet|abbreviat/i.test(k)) {
                            macroComponents.push({
                                propKey: k,
                                componentName: node.type?.name || '?',
                                valueType: typeof v,
                                valuePreview: JSON.stringify(v)?.substring(0, 200)
                            });
                        }
                    }
                }
                walkFiber(node.child, depth + 1);
                walkFiber(node.sibling, depth + 1);
            }
            walkFiber(rootEl[fiberKey], 0);
            results.fiberMacroComponents = macroComponents;
        }
    }

    // 3. Search all script tags for macro references
    const scripts = document.querySelectorAll('script[src]');
    const scriptUrls = [];
    scripts.forEach(s => {
        scriptUrls.push(s.src);
    });
    results.scriptSources = scriptUrls;

    // 4. Check for Service Workers
    results.hasServiceWorker = !!navigator.serviceWorker?.controller;

    // 5. Look for any context menus (right-click menus that might have macro insertion)
    const menus = document.querySelectorAll('[role="menu"], [role="menubar"], .MuiMenu-list, .MuiPopover-paper');
    results.menuCount = menus.length;
    const menuTexts = [];
    menus.forEach(m => {
        const items = m.querySelectorAll('[role="menuitem"], li');
        items.forEach(i => {
            menuTexts.push((i.innerText || '').trim().substring(0, 60));
        });
    });
    results.menuItems = menuTexts;

    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 13. TIPTAP/PROSEMIRROR DEEP DIVE"
Log ""
# ============================================================

$js = @'
(() => {
    const editors = document.querySelectorAll('.ProseMirror');
    if (editors.length === 0 || !editors[0].editor) return 'no editors';
    const editor = editors[0].editor;
    const results = {};

    // Full extension details
    results.extensions = editor.extensionManager.extensions.map(ext => {
        const info = {
            name: ext.name,
            type: ext.type,
            priority: ext.options?.priority
        };
        // Check for custom options
        if (ext.options && Object.keys(ext.options).length > 0) {
            info.options = {};
            for (const [k, v] of Object.entries(ext.options)) {
                const t = typeof v;
                if (t === 'function') info.options[k] = 'function';
                else if (t === 'object' && v !== null) info.options[k] = Object.keys(v).slice(0, 10);
                else info.options[k] = v;
            }
        }
        return info;
    });

    // Editor view properties
    results.viewProps = Object.keys(editor.view.props).sort();

    // Check for registered node views (custom rendering)
    results.nodeViews = Object.keys(editor.view.nodeViews || {});

    // Check options on the editor itself
    results.editorOptions = Object.keys(editor.options).sort();

    // Check for event handlers
    results.eventHandlers = [];
    if (editor.options.onUpdate) results.eventHandlers.push('onUpdate');
    if (editor.options.onSelectionUpdate) results.eventHandlers.push('onSelectionUpdate');
    if (editor.options.onCreate) results.eventHandlers.push('onCreate');
    if (editor.options.onDestroy) results.eventHandlers.push('onDestroy');
    if (editor.options.onTransaction) results.eventHandlers.push('onTransaction');
    if (editor.options.onFocus) results.eventHandlers.push('onFocus');
    if (editor.options.onBlur) results.eventHandlers.push('onBlur');

    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 14. ACCESSIBLE NAMES & ARIA TREE"
Log ""
# ============================================================

$js = @'
(() => {
    // Find all elements with aria-label, aria-describedby, data-testid, or role
    const items = [];
    document.querySelectorAll('[aria-label], [data-testid], [role]:not(option)').forEach(el => {
        const info = {
            tag: el.tagName,
            role: el.getAttribute('role'),
            ariaLabel: el.getAttribute('aria-label'),
            testId: el.getAttribute('data-testid'),
            id: el.id || undefined,
            text: (el.innerText || '').substring(0, 60).replace(/\n/g, ' '),
            visible: el.offsetParent !== null || el.tagName === 'BODY'
        };
        // Only include if visible or has interesting attributes
        if (info.visible || info.testId || info.ariaLabel) {
            items.push(info);
        }
    });
    return JSON.stringify(items, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# ============================================================
Log "## 15. MOSAIC-SPECIFIC: VOICE/DICTATION/DRAGON INTEGRATION POINTS"
Log ""
# ============================================================

$js = @'
(() => {
    const results = {};

    // Check for SpeechRecognition API
    results.hasSpeechRecognition = !!window.SpeechRecognition || !!window.webkitSpeechRecognition;

    // Check for WebRTC (might be used for audio streaming)
    results.hasRTCPeerConnection = !!window.RTCPeerConnection;

    // Check for MediaRecorder
    results.hasMediaRecorder = !!window.MediaRecorder;

    // Check for any audio-related elements
    results.audioElements = document.querySelectorAll('audio, video').length;

    // Check for WebSocket connections (Dragon might communicate via WS)
    // We can't enumerate existing connections, but check for patterns
    results.wsProto = !!window.WebSocket;

    // Look for dictation/voice related DOM elements
    const voiceEls = [];
    document.querySelectorAll('[class*="dictation"], [class*="voice"], [class*="speech"], [class*="record"], [class*="microphone"], [class*="mic"], [data-testid*="dictation"], [data-testid*="voice"], [aria-label*="dictation"], [aria-label*="voice"], [aria-label*="record"]').forEach(el => {
        voiceEls.push({
            tag: el.tagName,
            className: (el.className || '').substring(0, 100),
            ariaLabel: el.getAttribute('aria-label'),
            testId: el.getAttribute('data-testid'),
            text: (el.innerText || '').substring(0, 60)
        });
    });
    results.voiceElements = voiceEls;

    // Check for message event listeners (for cross-origin communication with Dragon)
    // Can't enumerate these, but check for postMessage patterns

    // Look for iframes within this iframe (Dragon widget?)
    const iframes = document.querySelectorAll('iframe');
    results.nestedIframes = [];
    iframes.forEach(f => {
        results.nestedIframes.push({
            src: f.src || f.getAttribute('srcdoc')?.substring(0, 100) || 'no-src',
            id: f.id || null,
            className: (f.className || '').substring(0, 100)
        });
    });

    return JSON.stringify(results, null, 2);
})()
'@
Log (Invoke-CDP $js)
Log ""

# Close WebSocket
$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', $ct).Wait()

# Save to file
$outputPath = 'C:\Users\erik.richter\Desktop\MosaicTools\CDP_INVENTORY.md'
$output.ToString() | Out-File -FilePath $outputPath -Encoding utf8
Write-Output "`n`n=== Results saved to $outputPath ==="

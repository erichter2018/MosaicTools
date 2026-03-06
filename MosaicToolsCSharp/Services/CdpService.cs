using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace MosaicTools.Services;

/// <summary>
/// Chrome DevTools Protocol client for Mosaic's WebView2.
/// Provides direct DOM access as an alternative to UI Automation.
/// Two WebSocket connections: SlimHub (patient metadata) and iframe (report editors).
/// Thread-safe synchronous public API — called from scrape timer thread.
/// </summary>
public class CdpService : IDisposable
{
    private const int CommandTimeoutMs = 3000;
    private const int ConnectThrottleMs = 10_000;
    private const int ClarioCdpPort = 9224;

    private int _port;
    private ClientWebSocket? _slimHubWs;
    private ClientWebSocket? _iframeWs;
    private string? _slimHubWsUrl;
    private string? _iframeWsUrl;
    private string? _lastIframeUrl; // Track iframe URL to detect report changes

    // Clario CDP (Chrome on port 9224)
    private ClientWebSocket? _clarioWs;
    private string? _clarioWsUrl;
    private CancellationTokenSource? _clarioCts;
    private readonly object _clarioSendLock = new();
    private long _lastClarioConnectAttemptTick64;

    private int _nextMessageId;
    private readonly ConcurrentDictionary<int, TaskCompletionSource<JsonElement>> _pending = new();
    private readonly object _slimHubSendLock = new();
    private readonly object _iframeSendLock = new();
    private CancellationTokenSource? _slimHubCts;
    private CancellationTokenSource? _iframeCts;
    private long _lastConnectAttemptTick64;
    private int _consecutiveTimeouts; // Track consecutive timeouts to auto-disconnect
    private string? _lastLoggedAccession; // Only log when scrape data changes
    private bool _firstScrapeLogged;
    private string? _lastKnownStudyType; // Cache: persists across ticks when input has focus
    private bool _studyTypeDiagLogged; // Only log DOM diagnostic once per empty streak
    private int _lastKnownFocusedEditor = 0; // Default to transcript (0), updated by scrape ticks
    private bool _mosaicMacrosFetched; // Only fetch once per iframe connection
    private bool _scrollFixActive;     // Track whether CSS scroll fix is injected in current iframe
    private double _columnRatio = 0.333; // Transcript:Report ratio (persisted via config)
    private string? _lastIframeHash;   // Smart scrape: skip full iframe scrape when hash unchanged
    private CdpScrapeResult? _lastCdpScrapeResult; // Cached result for hash-match reuse
    private bool _autoScrollWatcherActive; // Whether auto-scroll watcher interval is injected
    private bool _hideDragHandles = true; // Hide Tiptap drag handles in editor

    public bool IsConnected => _slimHubWs?.State == WebSocketState.Open;
    public bool IsIframeConnected => _iframeWs?.State == WebSocketState.Open;
    public bool IsClarioConnected => _clarioWs?.State == WebSocketState.Open;
    public bool ScrollFixActive => _scrollFixActive;
    public bool AutoScrollEnabled { get; set; } = true;
    public bool HideDragHandles { get => _hideDragHandles; set => _hideDragHandles = value; }

    /// <summary>Set column ratio from config on startup, and read back after drag.</summary>
    public double ColumnRatio
    {
        get => _columnRatio;
        set => _columnRatio = Math.Clamp(value, 0.15, 0.75);
    }

    /// <summary>
    /// Mosaic macros fetched from /macros/available API. Key = lowercase macro name, Value = expansion text.
    /// Populated on first scrape tick after iframe connects. Empty until fetched.
    /// </summary>
    public Dictionary<string, string> MosaicMacros { get; private set; } = new(StringComparer.OrdinalIgnoreCase);

    // ═══════ JS PAYLOADS ═══════

    // SlimHub: just return visible text — all parsing done in C#
    private const string JS_SCRAPE_SLIMHUB = @"(() => {
        return (document.body?.innerText || '').substring(0, 3000);
    })()";

    private const string JS_SCRAPE_IFRAME = @"(() => {
        const r = {};
        const active = document.activeElement;

        // Study type — try multiple selectors for robustness
        r.studyType = '';

        // Method 1: Known ID
        const studyInput = document.getElementById('studies-search');
        if (studyInput && active !== studyInput) {
            r.studyType = studyInput.value || '';
        }

        // Method 2: First combobox role with a non-empty input value
        if (!r.studyType) {
            const combos = document.querySelectorAll('[role=""combobox""]');
            for (const combo of combos) {
                const inp = (combo.tagName === 'INPUT') ? combo : combo.querySelector('input');
                if (inp && inp !== active && inp.value) {
                    r.studyType = inp.value;
                    break;
                }
            }
        }

        // Method 3: Input with aria-label containing 'study' or 'stud'
        if (!r.studyType) {
            for (const inp of document.querySelectorAll('input')) {
                const label = (inp.getAttribute('aria-label') || '') + (inp.getAttribute('placeholder') || '');
                if (/stud/i.test(label) && inp !== active && inp.value) {
                    r.studyType = inp.value;
                    break;
                }
            }
        }

        // Diagnostic when empty: report what we found so logs can reveal the DOM structure
        if (!r.studyType) {
            const combos = document.querySelectorAll('[role=""combobox""]');
            const comboVals = [];
            combos.forEach(c => {
                const inp = (c.tagName === 'INPUT') ? c : c.querySelector('input');
                comboVals.push((inp?.id || 'no-id') + '=' + JSON.stringify(inp?.value ?? null));
            });
            r.studyDiag = 'idExists=' + !!studyInput + ' combos=[' + comboVals.join(', ') + '] activeEl=' + (active?.tagName || '?') + '#' + (active?.id || '');
        }

        // Template name
        const templateInput = document.getElementById('templates-select');
        r.templateName = (templateInput && active !== templateInput) ? templateInput.value : '';

        // ProseMirror editors: [0]=transcript, [1]=final report
        // Custom DOM walker preserves <ol> numbering (CSS-generated, invisible to innerText)
        function walkEditor(node) {
            let t = '';
            for (const c of node.childNodes) {
                if (c.nodeType === 3) { t += c.textContent; continue; }
                const tag = c.tagName;
                if (tag === 'BR') { t += '\n'; continue; }
                if (tag === 'OL') {
                    let n = parseInt(c.getAttribute('start')) || 1;
                    for (const li of c.children) {
                        if (li.tagName === 'LI') { t += n + '. ' + walkEditor(li).trim() + '\n'; n++; }
                    }
                    continue;
                }
                if (tag === 'UL') {
                    for (const li of c.children) {
                        if (li.tagName === 'LI') { t += walkEditor(li).trim() + '\n'; }
                    }
                    continue;
                }
                if (tag === 'P' || tag === 'DIV' || (tag && tag.match(/^H[1-6]$/))) {
                    t += walkEditor(c) + '\n'; continue;
                }
                t += walkEditor(c);
            }
            return t;
        }
        const editors = document.querySelectorAll('.ProseMirror');
        r.editorCount = editors.length;
        if (editors.length >= 2) {
            r.reportText = walkEditor(editors[1]);
        } else if (editors.length >= 1) {
            r.reportText = walkEditor(editors[0]);
        }

        // Addendum detection
        if (r.reportText && r.reportText.trimStart().startsWith('Addendum')) {
            r.isAddendum = true;
        }

        // Track which editor has focus (for STT paste targeting)
        r.focusedEditor = -1;
        for (let i = 0; i < editors.length; i++) {
            if (editors[i].classList.contains('ProseMirror-focused')
                || editors[i].contains(document.activeElement)
                || editors[i] === document.activeElement) {
                r.focusedEditor = i; break;
            }
        }

        return JSON.stringify(r);
    })()";

    // Quick hash for smart scrape: lightweight fingerprint of iframe state.
    // If hash matches previous tick, skip the expensive JS_SCRAPE_IFRAME evaluation.
    private const string JS_QUICK_HASH = @"(() => {
        const editors = document.querySelectorAll('.ProseMirror');
        if (editors.length < 2) return 'E' + editors.length;
        const txt = editors[1].innerText || '';
        const len = txt.length;
        const h30 = txt.substring(0, 30);
        const t30 = len > 30 ? txt.substring(len - 30) : '';
        let focused = -1;
        for (let i = 0; i < editors.length; i++) {
            if (editors[i].classList.contains('ProseMirror-focused')
                || editors[i].contains(document.activeElement)) { focused = i; break; }
        }
        const studyInput = document.getElementById('studies-search');
        const sv = (studyInput && studyInput !== document.activeElement) ? studyInput.value : '';
        const drafted = document.body.innerText.includes('DRAFTED') ? 'D' : 'U';
        return len + '|' + h30 + '|' + t30 + '|' + focused + '|' + sv + '|' + drafted;
    })()";

    // Inject CSS to make Mosaic's three columns independently scrollable,
    // plus a draggable resize handle between transcript (col 0) and report (col 1).
    // Always removes old injection and re-discovers DOM (no stale persistence across reports).
    // {0} placeholder = column ratio (e.g. "0.333"), injected from C# at call time.
    private const string JS_INJECT_SCROLL_FIX_TEMPLATE = @"(() => {{
        // ── CLEANUP: remove previous injection + data attributes + stale inline styles ──
        document.getElementById('mt-scroll-fix')?.remove();
        document.querySelectorAll('[data-mt-resize-handle]').forEach(h => h.remove());
        const mtAttrs = ['data-mt-cols','data-mt-horizontal','data-mt-vertical','data-mt-col',
            'data-mt-col-editor','data-mt-col-scroll','data-mt-editor-wrapper','data-mt-editor-area',
            'data-mt-editor-inner','data-mt-scroll-area','data-mt-vr'];
        for (const attr of mtAttrs) {{
            document.querySelectorAll('[' + attr + ']').forEach(el => {{
                el.removeAttribute(attr);
                el.style.width = '';
                el.style.height = '';
                el.style.flex = '';
            }});
        }}
        // Also clear stale inline styles on ProseMirror ancestor chains — catches elements
        // that lost data-mt attrs when Mosaic's React re-rendered (e.g. orientation switch)
        document.querySelectorAll('.ProseMirror').forEach(editor => {{
            let el = editor.parentElement;
            while (el && el !== document.body) {{
                el.style.width = '';
                el.style.height = '';
                el.style.flex = '';
                el = el.parentElement;
            }}
        }});

        // ── DETECT: find editors and column structure ──
        const editors = document.querySelectorAll('.ProseMirror');
        if (editors.length < 2) return JSON.stringify({{error:'need_2_editors', found:editors.length}});

        function getAncestors(el) {{
            const a = [];
            while (el && el !== document.body) {{ a.push(el); el = el.parentElement; }}
            return a;
        }}
        const a0 = getAncestors(editors[0]);
        const s1 = new Set(getAncestors(editors[1]));
        let lca = null;
        for (const el of a0) {{ if (s1.has(el)) {{ lca = el; break; }} }}
        if (!lca) return JSON.stringify({{error:'no_lca'}});

        let colContainer = null;
        const lcaStyle = getComputedStyle(lca);
        if ((lcaStyle.display === 'flex' || lcaStyle.display === 'inline-flex') && lcaStyle.flexDirection === 'row') {{
            colContainer = lca;
        }} else {{
            let p = lca.parentElement;
            while (p && p !== document.body) {{
                const ps = getComputedStyle(p);
                if ((ps.display === 'flex' || ps.display === 'inline-flex') && ps.flexDirection === 'row') {{
                    const kids = Array.from(p.children).filter(c => getComputedStyle(c).display !== 'none' && c.offsetWidth > 80);
                    if (kids.length >= 2) {{ colContainer = p; break; }}
                }}
                p = p.parentElement;
            }}
        }}
        if (!colContainer) return JSON.stringify({{error:'no_flex_row_container'}});

        const columns = Array.from(colContainer.children).filter(c => {{
            const cs = getComputedStyle(c);
            return cs.display !== 'none' && c.offsetWidth > 50;
        }});
        if (columns.length < 2) return JSON.stringify({{error:'too_few_columns', found:columns.length}});

        const topPx = Math.round(colContainer.getBoundingClientRect().top);
        colContainer.setAttribute('data-mt-cols', '');

        // ── TAG: mark columns, wrappers, editor areas ──
        // Horizontal: col0=transcript, col1=report, col2=utility (2+ editor columns)
        // Vertical:   col0=stacked(transcript+report), col1=utility (1 editor column)
        let editorAreas = 0;
        let editorColCount = 0;
        columns.forEach((col, i) => {{
            col.setAttribute('data-mt-col', String(i));
            const editorsInCol = col.querySelectorAll('.ProseMirror');
            if (editorsInCol.length > 0) {{
                col.setAttribute('data-mt-col-editor', '');
                editorColCount++;
                editorsInCol.forEach(editor => {{
                    let wrapper = editor;
                    while (wrapper.parentElement && wrapper.parentElement !== col) {{
                        wrapper = wrapper.parentElement;
                    }}
                    if (wrapper && wrapper.parentElement === col) {{
                        wrapper.setAttribute('data-mt-editor-wrapper', '');
                        let inner = editor;
                        while (inner.parentElement && inner.parentElement !== wrapper) {{
                            inner = inner.parentElement;
                        }}
                        if (inner && inner.parentElement === wrapper) {{
                            inner.setAttribute('data-mt-editor-area', '');
                            editorAreas++;
                            // Mark ProseMirror's parent as scroll target (RichTextContent level)
                            if (editor.parentElement) {{
                                editor.parentElement.setAttribute('data-mt-scroll-area', '');
                            }}
                            let walk = editor.parentElement;
                            while (walk && walk !== inner) {{
                                walk.setAttribute('data-mt-editor-inner', '');
                                walk = walk.parentElement;
                            }}
                        }}
                    }}
                }});
            }} else {{
                col.setAttribute('data-mt-col-scroll', '');
            }}
        }});

        // ── LAYOUT: detect orientation, insert resize handle, compute dynamic CSS values ──
        const isHorizontal = editorColCount >= 2;
        colContainer.setAttribute(isHorizontal ? 'data-mt-horizontal' : 'data-mt-vertical', '');

        // Dynamic CSS values — embedded in style tag, NOT set as inline styles.
        // This ensures orientation switches leave no stale inline styles behind.
        let col0Width = 0, col1Width = 0, vRatio = 0.3;

        if (isHorizontal) {{
            const w0 = columns[0].offsetWidth;
            const w1 = columns[1].offsetWidth;

            const handle = document.createElement('div');
            handle.setAttribute('data-mt-resize-handle', '');
            colContainer.insertBefore(handle, columns[1]);

            const ratio = window.__mtColumnRatio || {0};
            const totalW = w0 + w1 - 6;
            col0Width = Math.round(ratio * totalW);
            col1Width = totalW - col0Width;
            window.__mtColumnRatio = ratio;

            // Drag: uses inline setProperty('important') for real-time updates only
            handle.addEventListener('mousedown', e => {{
                e.preventDefault();
                handle.classList.add('dragging');
                const tw = columns[0].offsetWidth + columns[1].offsetWidth;
                const startOffset = columns[0].getBoundingClientRect().left;
                function onMove(ev) {{
                    const newCol0Width = ev.clientX - startOffset;
                    let r = newCol0Width / tw;
                    r = Math.max(0.15, Math.min(0.75, r));
                    const cw0 = Math.round(r * tw);
                    columns[0].style.setProperty('width', cw0 + 'px', 'important');
                    columns[1].style.setProperty('width', (tw - cw0) + 'px', 'important');
                    window.__mtColumnRatio = r;
                }}
                function onUp() {{
                    handle.classList.remove('dragging');
                    document.removeEventListener('mousemove', onMove);
                    document.removeEventListener('mouseup', onUp);
                }}
                document.addEventListener('mousemove', onMove);
                document.addEventListener('mouseup', onUp);
            }});
        }} else {{
            // Vertical layout: resize handle between stacked wrappers
            const editorCol = columns.find(c => c.hasAttribute('data-mt-col-editor'));
            if (editorCol) {{
                const wrappers = Array.from(editorCol.children).filter(c => c.hasAttribute('data-mt-editor-wrapper'));
                if (wrappers.length >= 2) {{
                    wrappers[0].setAttribute('data-mt-vr', '0');
                    wrappers[1].setAttribute('data-mt-vr', '1');

                    const handle = document.createElement('div');
                    handle.setAttribute('data-mt-resize-handle', '');
                    handle.setAttribute('data-mt-resize-horizontal', '');
                    editorCol.insertBefore(handle, wrappers[1]);

                    vRatio = window.__mtVerticalRatio || 0.3;
                    window.__mtVerticalRatio = vRatio;

                    handle.addEventListener('mousedown', e => {{
                        e.preventDefault();
                        handle.classList.add('dragging');
                        const th = wrappers[0].offsetHeight + wrappers[1].offsetHeight;
                        const startTop = wrappers[0].getBoundingClientRect().top;
                        function onMove(ev) {{
                            let r = (ev.clientY - startTop) / th;
                            r = Math.max(0.1, Math.min(0.85, r));
                            wrappers[0].style.setProperty('flex', r + ' 0 0px', 'important');
                            wrappers[1].style.setProperty('flex', (1 - r) + ' 0 0px', 'important');
                            window.__mtVerticalRatio = r;
                        }}
                        function onUp() {{
                            handle.classList.remove('dragging');
                            document.removeEventListener('mousemove', onMove);
                            document.removeEventListener('mouseup', onUp);
                        }}
                        document.addEventListener('mousemove', onMove);
                        document.addEventListener('mouseup', onUp);
                    }});
                }}
            }}
        }}

        // ── CSS: all layout values embedded in stylesheet, no inline styles ──
        const hideDragHandles = {1};
        const styleEl = document.createElement('style');
        styleEl.id = 'mt-scroll-fix';
        let css = `/* MosaicTools: independent column scrolling + resize */
html, body {{ overflow: hidden !important; }}
[data-mt-cols] {{
    height: calc(100vh - ${{topPx}}px) !important;
    max-height: calc(100vh - ${{topPx}}px) !important;
    overflow: hidden !important;
}}
[data-mt-col] {{
    height: 100% !important;
    overflow: hidden !important;
}}
[data-mt-col-editor] {{
    display: flex !important;
    flex-direction: column !important;
    overflow: hidden !important;
}}
[data-mt-col-scroll] {{
    overflow-y: auto !important;
}}
[data-mt-editor-wrapper] {{
    display: flex !important;
    flex-direction: column !important;
    flex-wrap: nowrap !important;
    overflow: hidden !important;
    min-height: 0 !important;
}}
[data-mt-editor-wrapper] > *:not([data-mt-editor-area]) {{
    flex: 0 0 auto !important;
}}
[data-mt-editor-area] {{
    flex: 1 1 0 !important;
    overflow-y: auto !important;
    min-height: 0 !important;
    max-height: none !important;
}}
[data-mt-editor-inner] {{
    overflow: visible !important;
}}
[data-mt-editor-area] .ProseMirror {{
    overflow: visible !important;
}}
/* ── Horizontal mode ── */
[data-mt-horizontal] > [data-mt-col-editor] {{
    box-sizing: border-box !important;
    max-width: none !important;
    flex: 0 0 auto !important;
}}
[data-mt-horizontal] [data-mt-editor-wrapper] {{
    flex: 1 1 0 !important;
}}
[data-mt-horizontal] > [data-mt-col=""0""][data-mt-col-editor] {{
    width: ${{col0Width}}px !important;
}}
[data-mt-horizontal] > [data-mt-col=""1""][data-mt-col-editor] {{
    width: ${{col1Width}}px !important;
}}
/* ── Vertical mode ── */
[data-mt-vertical] > [data-mt-col-editor] {{
    max-width: none !important;
    flex: 1 1 0 !important;
    min-width: 0 !important;
    overflow: hidden !important;
}}
[data-mt-vertical] [data-mt-editor-wrapper] {{
    max-width: 100% !important;
    min-width: 0 !important;
}}
[data-mt-vertical] [data-mt-editor-area] {{
    flex-wrap: nowrap !important;
    overflow: hidden !important;
}}
[data-mt-vertical] [data-mt-editor-area] > *:not([data-mt-editor-inner]) {{
    flex: 0 0 auto !important;
}}
[data-mt-vertical] [data-mt-editor-inner] {{
    display: flex !important;
    flex-direction: column !important;
    flex: 1 1 0 !important;
    height: auto !important;
    min-height: 0 !important;
    max-height: none !important;
    overflow: hidden !important;
}}
[data-mt-vertical] [data-mt-editor-inner] > *:not([data-mt-editor-inner]):not(.ProseMirror):not([data-mt-scroll-area]) {{
    flex: 0 0 auto !important;
}}
[data-mt-vertical] [data-mt-scroll-area] {{
    flex: 1 1 0 !important;
    min-height: 0 !important;
    overflow-y: auto !important;
}}
[data-mt-vertical] [data-mt-editor-area] .ProseMirror {{
    overflow: visible !important;
}}
[data-mt-vertical] [data-mt-editor-wrapper] [class*=""MuiGrid""] {{
    max-width: 100% !important;
}}
[data-mt-vertical] [data-mt-editor-wrapper][data-mt-vr=""0""] {{
    flex: ${{vRatio}} 0 0px !important;
    min-height: 0 !important;
}}
[data-mt-vertical] [data-mt-editor-wrapper][data-mt-vr=""1""] {{
    flex: ${{1 - vRatio}} 0 0px !important;
    min-height: 0 !important;
}}
/* ── Resize handle ── */
[data-mt-resize-handle] {{
    width: 6px; cursor: col-resize;
    background: rgba(255, 255, 255, 0.06);
    position: relative; z-index: 10; flex-shrink: 0;
    transition: background 0.15s;
}}
[data-mt-resize-handle]::after {{
    content: '';
    position: absolute;
    left: 2px; top: 50%; transform: translateY(-50%);
    width: 2px; height: 32px;
    background: rgba(255, 255, 255, 0.15);
    border-radius: 1px;
}}
[data-mt-resize-horizontal] {{
    width: auto !important; height: 6px !important;
    cursor: row-resize !important;
}}
[data-mt-resize-horizontal]::after {{
    left: 50% !important; top: 2px !important;
    transform: translateX(-50%) !important;
    width: 32px !important; height: 2px !important;
}}
[data-mt-resize-handle]:hover, [data-mt-resize-handle].dragging {{
    background: rgba(100, 150, 255, 0.3);
}}
[data-mt-resize-handle]:hover::after, [data-mt-resize-handle].dragging::after {{
    background: rgba(100, 150, 255, 0.6);
}}`;
        if (hideDragHandles) {{
            css += `
.ProseMirror [data-drag-handle] {{ display: none !important; }}
.ProseMirror [data-testid=""DeleteIcon""] {{ display: none !important; }}
.ProseMirror button:has(> [data-testid=""DeleteIcon""]) {{ display: none !important; }}
.ProseMirror button:has(> [data-testid=""DragHandleIcon""]) {{ display: none !important; }}`;
        }}
        styleEl.textContent = css;
        document.head.appendChild(styleEl);

        // ── DIAGNOSTICS ──
        const wrapperInfo = [];
        document.querySelectorAll('[data-mt-editor-wrapper]').forEach(w => {{
            const a = w.querySelector('[data-mt-editor-area]');
            const p = w.querySelector('.ProseMirror');
            const wr = w.getBoundingClientRect();
            const pr = p ? p.getBoundingClientRect() : {{left:0,top:0,width:0}};
            wrapperInfo.push({{ww: w.offsetWidth, wh: w.offsetHeight, aw: a ? a.offsetWidth : 0, ah: a ? a.offsetHeight : 0,
                wl: Math.round(wr.left), wt: Math.round(wr.top), pl: Math.round(pr.left), pt: Math.round(pr.top), pw: Math.round(pr.width)}});
        }});
        return JSON.stringify({{
            ok: true,
            layout: isHorizontal ? 'horizontal' : 'vertical',
            columns: columns.length,
            topPx: topPx,
            editorAreas: editorAreas,
            editorCols: editorColCount,
            ratio: isHorizontal ? window.__mtColumnRatio : window.__mtVerticalRatio,
            hideDragHandles: hideDragHandles,
            wrappers: wrapperInfo,
            containerW: colContainer.offsetWidth,
            windowW: window.innerWidth,
            colWidths: columns.map(c => c.offsetWidth),
            containerTag: colContainer.tagName,
            containerClass: (colContainer.className || '').substring(0, 100)
        }});
    }})()";

    private const string JS_REMOVE_SCROLL_FIX = @"(() => {
        document.getElementById('mt-scroll-fix')?.remove();
        document.querySelectorAll('[data-mt-resize-handle]').forEach(h => h.remove());
        const attrs = ['data-mt-cols','data-mt-horizontal','data-mt-vertical','data-mt-col',
            'data-mt-col-editor','data-mt-col-scroll','data-mt-editor-wrapper','data-mt-editor-area',
            'data-mt-editor-inner','data-mt-scroll-area','data-mt-vr'];
        for (const attr of attrs) {
            document.querySelectorAll('[' + attr + ']').forEach(el => {
                el.removeAttribute(attr);
                el.style.width = '';
                el.style.height = '';
                el.style.flex = '';
            });
        }
        document.querySelectorAll('.ProseMirror').forEach(editor => {
            let el = editor.parentElement;
            while (el && el !== document.body) {
                el.style.width = '';
                el.style.height = '';
                el.style.flex = '';
                el = el.parentElement;
            }
        });
        return 'removed';
    })()";

    // ═══════ CONNECTION MANAGEMENT ═══════

    /// <summary>
    /// Set the WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS env var to enable remote debugging.
    /// Must be set before Mosaic starts — requires Mosaic restart.
    /// </summary>
    public static void EnsureEnvVar()
    {
        const string varName = "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS";
        const string required = "--remote-debugging-port=0";
        var current = Environment.GetEnvironmentVariable(varName, EnvironmentVariableTarget.User);
        if (current != null && current.Contains("--remote-debugging-port"))
            return;

        var newValue = string.IsNullOrEmpty(current) ? required : $"{current} {required}";
        Environment.SetEnvironmentVariable(varName, newValue, EnvironmentVariableTarget.User);
        Logger.Trace($"CDP: Set {varName}={newValue} (user-level). Mosaic restart required.");
    }

    /// <summary>
    /// Discover the DevTools port from the DevToolsActivePort file, find targets, connect WebSockets.
    /// Returns true if at least SlimHub is connected.
    /// </summary>
    public bool TryConnect()
    {
        long now = Environment.TickCount64;
        if (now - _lastConnectAttemptTick64 < ConnectThrottleMs)
            return IsConnected;
        _lastConnectAttemptTick64 = now;

        try
        {
            if (!DiscoverDevToolsPort())
            {
                Logger.Trace("CDP: DevToolsActivePort file not found");
                return false;
            }

            DiscoverAndConnect();
            return IsConnected;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: TryConnect failed: {ex.Message}");
            return false;
        }
    }

    private bool DiscoverDevToolsPort()
    {
        // Find MosaicInfoHub package directory
        var packagesDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Packages");

        if (!Directory.Exists(packagesDir))
            return false;

        var mosaicDirs = Directory.GetDirectories(packagesDir, "MosaicInfoHub_*");
        if (mosaicDirs.Length == 0)
            return false;

        var portFile = Path.Combine(mosaicDirs[0], "LocalState", "EBWebView", "DevToolsActivePort");
        if (!File.Exists(portFile))
            return false;

        var lines = File.ReadAllLines(portFile);
        if (lines.Length == 0 || !int.TryParse(lines[0].Trim(), out var port))
            return false;

        _port = port;
        return true;
    }

    private void DiscoverAndConnect()
    {
        // HTTP GET target list
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
        var json = http.GetStringAsync($"http://localhost:{_port}/json").GetAwaiter().GetResult();
        var targets = JsonSerializer.Deserialize<JsonElement>(json);

        string? slimHubWs = null;
        string? iframeWs = null;
        string? iframeUrl = null;

        foreach (var target in targets.EnumerateArray())
        {
            var url = target.TryGetProperty("url", out var urlEl) ? urlEl.GetString() ?? "" : "";
            var wsUrl = target.TryGetProperty("webSocketDebuggerUrl", out var wsEl) ? wsEl.GetString() : null;
            var title = target.TryGetProperty("title", out var titleEl) ? titleEl.GetString() : "";
            var type = target.TryGetProperty("type", out var t) ? t.GetString() : "";

            // Match SlimHub by title (most reliable) or URL pattern
            if ((title == "SlimHub" || url.Contains("/index.html#/reporting")) && type == "page")
            {
                if (slimHubWs == null) slimHubWs = wsUrl;
            }
            // Match iframe by URL — type can be "page" or "iframe"
            else if (url.Contains("rp.radpair.com") && url.Contains("/reports/"))
            {
                iframeWs = wsUrl;
                iframeUrl = url;
            }
        }

        // Connect SlimHub if not already connected
        if (!string.IsNullOrEmpty(slimHubWs) && !IsConnected)
        {
            _slimHubWsUrl = slimHubWs;
            Logger.Trace($"CDP: Connecting to SlimHub: {slimHubWs}");
            ConnectSlimHub();
        }
        else if (slimHubWs == null && !IsConnected)
        {
            Logger.Trace("CDP: SlimHub target not found or no webSocketDebuggerUrl");
        }

        // Connect iframe (reconnect if URL changed = new report)
        if (!string.IsNullOrEmpty(iframeWs) && (iframeWs != _iframeWsUrl || !IsIframeConnected))
        {
            DisconnectIframe();
            _iframeWsUrl = iframeWs;
            _lastIframeUrl = iframeUrl;
            Logger.Trace($"CDP: Connecting to iframe: {iframeUrl}");
            ConnectIframe();
        }
    }

    private void ConnectSlimHub()
    {
        try
        {
            _slimHubCts?.Cancel();
            _slimHubCts = new CancellationTokenSource();
            _slimHubWs = new ClientWebSocket();
            _slimHubWs.ConnectAsync(new Uri(_slimHubWsUrl!), CancellationToken.None).GetAwaiter().GetResult();

            // Dedicated receive thread (not async Task) — avoids thread pool scheduling issues
            var ws = _slimHubWs;
            var ct = _slimHubCts.Token;
            var thread = new Thread(() => ReceiveLoopSync(ws, ct, "SlimHub"))
            {
                IsBackground = true,
                Name = "CDP-SlimHub-Recv"
            };
            thread.Start();

            _consecutiveTimeouts = 0;
            Logger.Trace($"CDP: Connected to SlimHub (port {_port}, wsUrl={_slimHubWsUrl})");
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: SlimHub connect failed: {ex.Message}");
            _slimHubWs = null;
        }
    }

    private void ConnectIframe()
    {
        try
        {
            _iframeCts?.Cancel();
            _iframeCts = new CancellationTokenSource();
            _iframeWs = new ClientWebSocket();
            _iframeWs.ConnectAsync(new Uri(_iframeWsUrl!), CancellationToken.None).GetAwaiter().GetResult();

            var ws = _iframeWs;
            var ct = _iframeCts.Token;
            var thread = new Thread(() => ReceiveLoopSync(ws, ct, "Iframe"))
            {
                IsBackground = true,
                Name = "CDP-Iframe-Recv"
            };
            thread.Start();

            Logger.Trace($"CDP: Connected to iframe ({_lastIframeUrl})");
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: Iframe connect failed: {ex.Message}");
            _iframeWs = null;
        }
    }

    private void DisconnectIframe()
    {
        try
        {
            _iframeCts?.Cancel();
            if (_iframeWs?.State == WebSocketState.Open)
                _iframeWs.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None)
                    .GetAwaiter().GetResult();
        }
        catch { }
        _iframeWs = null;
        _iframeWsUrl = null;
        _mosaicMacrosFetched = false; // Re-fetch macros on next connection
        _scrollFixActive = false;     // Re-inject scroll fix on next connection
        _lastIframeHash = null;       // Reset hash cache on disconnect
        _lastCdpScrapeResult = null;
        _autoScrollWatcherActive = false;
    }

    private void DisconnectSlimHub()
    {
        try
        {
            _slimHubCts?.Cancel();
            if (_slimHubWs?.State == WebSocketState.Open)
                _slimHubWs.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None)
                    .GetAwaiter().GetResult();
        }
        catch { }
        _slimHubWs = null;
        _slimHubWsUrl = null;
    }

    // ═══════ CLARIO CDP CONNECTION (Chrome on port 9224) ═══════

    /// <summary>
    /// Try to connect to Clario's Chrome instance via CDP on port 9224.
    /// Throttled to one attempt per 10 seconds.
    /// </summary>
    public bool TryConnectClario()
    {
        long now = Environment.TickCount64;
        if (now - _lastClarioConnectAttemptTick64 < ConnectThrottleMs)
            return IsClarioConnected;
        _lastClarioConnectAttemptTick64 = now;

        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
            var json = http.GetStringAsync($"http://127.0.0.1:{ClarioCdpPort}/json/list").GetAwaiter().GetResult();
            var targets = JsonSerializer.Deserialize<JsonElement>(json);

            string? clarioWs = null;
            foreach (var target in targets.EnumerateArray())
            {
                var title = target.TryGetProperty("title", out var titleEl) ? titleEl.GetString() ?? "" : "";
                var wsUrl = target.TryGetProperty("webSocketDebuggerUrl", out var wsEl) ? wsEl.GetString() : null;

                if (title.Contains("Clario", StringComparison.OrdinalIgnoreCase) &&
                    title.Contains("Worklist", StringComparison.OrdinalIgnoreCase))
                {
                    clarioWs = wsUrl;
                    break;
                }
            }

            if (!string.IsNullOrEmpty(clarioWs) && (clarioWs != _clarioWsUrl || !IsClarioConnected))
            {
                DisconnectClario();
                _clarioWsUrl = clarioWs;
                Logger.Trace($"CDP: Connecting to Clario: {clarioWs}");
                ConnectClario();
            }
            else if (clarioWs == null && !IsClarioConnected)
            {
                Logger.Trace("CDP: Clario target not found on port 9224");
            }

            return IsClarioConnected;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: TryConnectClario failed: {ex.Message}");
            return false;
        }
    }

    private void ConnectClario()
    {
        try
        {
            _clarioCts?.Cancel();
            _clarioCts = new CancellationTokenSource();
            _clarioWs = new ClientWebSocket();
            _clarioWs.ConnectAsync(new Uri(_clarioWsUrl!), CancellationToken.None).GetAwaiter().GetResult();

            var ws = _clarioWs;
            var ct = _clarioCts.Token;
            var thread = new Thread(() => ReceiveLoopSync(ws, ct, "Clario"))
            {
                IsBackground = true,
                Name = "CDP-Clario-Recv"
            };
            thread.Start();

            Logger.Trace($"CDP: Connected to Clario (port {ClarioCdpPort})");
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: Clario connect failed: {ex.Message}");
            _clarioWs = null;
        }
    }

    private void DisconnectClario()
    {
        try
        {
            _clarioCts?.Cancel();
            if (_clarioWs?.State == WebSocketState.Open)
                _clarioWs.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None)
                    .GetAwaiter().GetResult();
        }
        catch { }
        _clarioWs = null;
        _clarioWsUrl = null;
    }

    // ═══════ WEBSOCKET COMMUNICATION ═══════

    private JsonElement? SendCommand(ClientWebSocket? ws, object sendLock, string js, int timeoutMs = CommandTimeoutMs)
    {
        if (ws?.State != WebSocketState.Open) return null;

        int id = Interlocked.Increment(ref _nextMessageId);
        var tcs = new TaskCompletionSource<JsonElement>(TaskCreationOptions.RunContinuationsAsynchronously);
        _pending[id] = tcs;

        try
        {
            var payload = JsonSerializer.Serialize(new
            {
                id,
                method = "Runtime.evaluate",
                @params = new { expression = js, returnByValue = true }
            });

            var bytes = Encoding.UTF8.GetBytes(payload);
            lock (sendLock)
            {
                ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }

            if (tcs.Task.Wait(timeoutMs))
            {
                _consecutiveTimeouts = 0;
                return tcs.Task.Result;
            }

            _consecutiveTimeouts++;
            Logger.Trace($"CDP: Command {id} timed out ({timeoutMs}ms), consecutive={_consecutiveTimeouts}");

            // After 3 consecutive timeouts, disconnect — receive loop is dead
            if (_consecutiveTimeouts >= 3)
            {
                Logger.Trace("CDP: Too many consecutive timeouts, disconnecting");
                DisconnectSlimHub();
                DisconnectIframe();
            }

            return null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: SendCommand error: {ex.Message}");
            return null;
        }
        finally
        {
            _pending.TryRemove(id, out _);
        }
    }

    private JsonElement? SendToSlimHub(string js) => SendCommand(_slimHubWs, _slimHubSendLock, js);
    private JsonElement? SendToIframe(string js) => SendCommand(_iframeWs, _iframeSendLock, js);
    private JsonElement? SendToClario(string js) => SendCommand(_clarioWs, _clarioSendLock, js);

    /// <summary>Send async JS (returns a Promise) to iframe with awaitPromise, longer timeout.</summary>
    private string? SendToIframeAsync(string js)
    {
        if (_iframeWs?.State != WebSocketState.Open) return null;

        int id = Interlocked.Increment(ref _nextMessageId);
        var tcs = new TaskCompletionSource<JsonElement>(TaskCreationOptions.RunContinuationsAsynchronously);
        _pending[id] = tcs;

        try
        {
            var payload = JsonSerializer.Serialize(new
            {
                id,
                method = "Runtime.evaluate",
                @params = new { expression = js, returnByValue = true, awaitPromise = true }
            });

            var bytes = Encoding.UTF8.GetBytes(payload);
            lock (_iframeSendLock)
            {
                _iframeWs.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }

            // Longer timeout for network fetch (10s)
            if (tcs.Task.Wait(10_000))
                return ExtractResultValue(tcs.Task.Result);

            Logger.Trace($"CDP: Async command {id} timed out (10s)");
            return null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: SendToIframeAsync error: {ex.Message}");
            return null;
        }
        finally
        {
            _pending.TryRemove(id, out _);
        }
    }

    /// <summary>
    /// Synchronous receive loop running on a dedicated background thread.
    /// Reads WebSocket messages and completes pending TaskCompletionSources.
    /// </summary>
    private void ReceiveLoopSync(ClientWebSocket ws, CancellationToken ct, string label)
    {
        var buf = new byte[256 * 1024];
        Logger.Trace($"CDP: {label} receive loop started (thread {Environment.CurrentManagedThreadId})");
        try
        {
            while (!ct.IsCancellationRequested && ws.State == WebSocketState.Open)
            {
                var ms = new MemoryStream();
                WebSocketReceiveResult result;
                do
                {
                    var task = ws.ReceiveAsync(new ArraySegment<byte>(buf), ct);
                    task.Wait(ct); // Block this thread until data arrives
                    result = task.Result;
                    ms.Write(buf, 0, result.Count);
                } while (!result.EndOfMessage);

                if (result.MessageType == WebSocketMessageType.Close)
                {
                    Logger.Trace($"CDP: {label} WebSocket closed by server");
                    break;
                }

                var json = Encoding.UTF8.GetString(ms.GetBuffer(), 0, (int)ms.Length);
                try
                {
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("id", out var idEl))
                    {
                        int id = idEl.GetInt32();
                        if (_pending.TryRemove(id, out var tcs))
                        {
                            tcs.TrySetResult(doc.RootElement.Clone());
                        }
                    }
                    // else: event/notification without id — ignore
                }
                catch (Exception ex)
                {
                    Logger.Trace($"CDP: {label} parse error: {ex.Message}, json={json.Substring(0, Math.Min(json.Length, 200))}");
                }
            }
        }
        catch (OperationCanceledException) { }
        catch (AggregateException ae) when (ae.InnerException is OperationCanceledException) { }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: {label} receive loop error: {ex.GetType().Name}: {ex.Message}");
        }
        Logger.Trace($"CDP: {label} receive loop ended (ws.State={ws.State})");
    }

    /// <summary>
    /// Extract the string value from a Runtime.evaluate response.
    /// Response shape: { "result": { "result": { "type": "string", "value": "..." } } }
    /// </summary>
    private string? ExtractResultValue(JsonElement? response)
    {
        if (response == null) return null;
        try
        {
            return response.Value
                .GetProperty("result")
                .GetProperty("result")
                .GetProperty("value")
                .GetString();
        }
        catch { return null; }
    }

    // ═══════ PUBLIC API: SCRAPE ═══════

    /// <summary>
    /// Scrape all metadata from SlimHub + report text from iframe in one call.
    /// Returns null if SlimHub is not connected or command times out.
    /// </summary>
    public CdpScrapeResult? Scrape()
    {
        if (!IsConnected) return null;

        // If iframe disconnected, try to reconnect (throttled — same as main connect)
        if (!IsIframeConnected)
        {
            long now = Environment.TickCount64;
            if (now - _lastConnectAttemptTick64 >= ConnectThrottleMs)
            {
                _lastConnectAttemptTick64 = now;
                try { DiscoverAndConnect(); } catch { }
            }
        }

        // Fetch Mosaic macros on first scrape tick after iframe connects
        if (IsIframeConnected && !_mosaicMacrosFetched)
        {
            try { FetchMosaicMacros(); }
            catch (Exception ex) { Logger.Trace($"CDP: FetchMosaicMacros exception: {ex.Message}"); }
        }

        var result = new CdpScrapeResult();

        // Scrape SlimHub — JS returns raw visible text, all parsing done in C#
        var visibleText = ExtractResultValue(SendToSlimHub(JS_SCRAPE_SLIMHUB));
        if (visibleText == null)
            return null; // SlimHub not responding — fall through to UIA

        try
        {
            ParseVisibleText(result, visibleText);
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: SlimHub parse error: {ex.Message}");
            return null;
        }

        // Clear study-type cache on study change BEFORE iframe scrape
        // (prevents stale template name from previous study on first tick of new study)
        if (_lastLoggedAccession != null && result.Accession != _lastLoggedAccession)
        {
            _lastKnownStudyType = null;
            _mosaicMacrosFetched = false; // Re-fetch macros on study change
            _lastIframeHash = null;       // Force full scrape on study change
            _lastCdpScrapeResult = null;
        }

        // Scrape iframe (optional — iframe may not be connected if no report open)
        if (IsIframeConnected)
        {
            // Smart scrape: check lightweight hash first to skip expensive full scrape
            bool hashChanged = true;
            var quickHash = ExtractResultValue(SendToIframe(JS_QUICK_HASH));
            if (quickHash != null && quickHash == _lastIframeHash && _lastCdpScrapeResult != null)
            {
                // Hash unchanged — reuse cached iframe data
                hashChanged = false;
                result.TemplateName = _lastCdpScrapeResult.TemplateName;
                result.ReportText = _lastCdpScrapeResult.ReportText;
                result.IsAddendum = _lastCdpScrapeResult.IsAddendum;
            }

            if (hashChanged)
            {
                var iframeJson = ExtractResultValue(SendToIframe(JS_SCRAPE_IFRAME));
                if (iframeJson != null)
                {
                    try
                    {
                        var data = JsonSerializer.Deserialize<JsonElement>(iframeJson);
                        var studyType = GetStr(data, "studyType");
                        var templateName = GetStr(data, "templateName");

                        // Prefer studyType (exam type selector), fall back to templateName
                        var effectiveType = !string.IsNullOrEmpty(studyType) ? studyType
                            : !string.IsNullOrEmpty(templateName) ? templateName : null;

                        if (!string.IsNullOrEmpty(effectiveType))
                            _lastKnownStudyType = effectiveType;
                        result.TemplateName = _lastKnownStudyType;
                        result.ReportText = GetStr(data, "reportText");
                        result.IsAddendum = GetBool(data, "isAddendum");

                        // Log diagnostic once when study type can't be read from DOM
                        if (string.IsNullOrEmpty(effectiveType) && !_studyTypeDiagLogged)
                        {
                            _studyTypeDiagLogged = true;
                            var diag = GetStr(data, "studyDiag");
                            Logger.Trace($"CDP: StudyType empty from DOM — {diag ?? "no diag"}");
                        }
                        else if (!string.IsNullOrEmpty(effectiveType))
                        {
                            _studyTypeDiagLogged = false; // Reset so we log again if it breaks
                        }

                        // Track focused editor for STT paste targeting
                        if (data.TryGetProperty("focusedEditor", out var fe)
                            && fe.ValueKind == JsonValueKind.Number)
                        {
                            int focused = fe.GetInt32();
                            if (focused >= 0) _lastKnownFocusedEditor = focused;
                        }

                        // Update hash cache
                        _lastIframeHash = quickHash;
                    }
                    catch (Exception ex)
                    {
                        Logger.Trace($"CDP: Iframe parse error: {ex.Message}");
                    }
                }
            }
        }

        // Cache result for hash-match reuse next tick
        _lastCdpScrapeResult = result;

        // Log on first scrape and when accession changes (avoid spamming every tick)
        if (!_firstScrapeLogged || result.Accession != _lastLoggedAccession)
        {
            _firstScrapeLogged = true;
            _lastLoggedAccession = result.Accession;
            Logger.Trace($"CDP: Scraped → acc={result.Accession}, name={result.PatientName}, mrn={result.Mrn}, site={result.SiteCode}, desc={result.Description}, tmpl={result.TemplateName}, drafted={result.IsDrafted}, addendum={result.IsAddendum}, report={result.ReportText?.Length ?? 0} chars");
        }
        return result;
    }

    /// <summary>
    /// Parse the visible text from SlimHub body to extract all patient/study metadata.
    /// Text structure (line by line):
    ///   WILLIAMS LATRINA                     ← patient name (first all-caps line)
    ///   FEMALE, AGE 45, DOB: 05/03/1980      ← gender, age
    ///   Study Date: 03/02/2026
    ///   Body Part: ABDOMEN,PELVIS
    ///   MRN: 2132099CRE                      ← MRN
    ///   Ordering: RANGINWALA ADAM
    ///   Site Group: CIRPA
    ///   Site Code: CRE                       ← site code
    ///   Description: CT ABDOMEN PELVIS ...    ← description
    ///   Reason for visit: ...
    ///   Current Study
    ///   UNDRAFTED / DRAFTED                   ← draft status
    ///   317276220260302CRE                    ← accession (last line, digits + optional site code)
    /// </summary>
    private static void ParseVisibleText(CdpScrapeResult result, string? visibleText)
    {
        if (string.IsNullOrEmpty(visibleText)) return;
        var lines = visibleText.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (line.Length == 0) continue;

            // Patient name: first all-caps alphabetic line (e.g. "SMITH, JOHN" or "WILLIAMS LATRINA")
            if (result.PatientName == null && line.Length > 2 && line == line.ToUpperInvariant()
                && System.Text.RegularExpressions.Regex.IsMatch(line, @"^[A-Z][A-Z ,'\-]+$"))
            {
                result.PatientName = line;
                continue;
            }

            // Gender/Age: "FEMALE, AGE 45, DOB: ..." or "MALE, AGE 72"
            if (result.PatientGender == null)
            {
                var ageSexMatch = System.Text.RegularExpressions.Regex.Match(line, @"\b(FEMALE|MALE)\b.*?\bAGE\s*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (ageSexMatch.Success)
                {
                    result.PatientGender = ageSexMatch.Groups[1].Value.StartsWith("F", StringComparison.OrdinalIgnoreCase) ? "Female" : "Male";
                    result.PatientAge = int.Parse(ageSexMatch.Groups[2].Value);
                    continue;
                }
            }

            // Labeled fields (case-insensitive prefix match)
            if (line.StartsWith("MRN:", StringComparison.OrdinalIgnoreCase))
            {
                result.Mrn = line.Substring(4).Trim();
                continue;
            }
            if (line.StartsWith("Site Code:", StringComparison.OrdinalIgnoreCase))
            {
                result.SiteCode = line.Substring("Site Code:".Length).Trim();
                continue;
            }
            if (line.StartsWith("Description:", StringComparison.OrdinalIgnoreCase))
            {
                result.Description = line.Substring("Description:".Length).Trim();
                continue;
            }

            // DRAFTED: if "DRAFTED" appears (standalone), study IS drafted — takes priority
            if (line == "DRAFTED")
            {
                result.IsDrafted = true;
                continue;
            }
            // UNDRAFTED: only set false if we haven't already seen DRAFTED
            if (line == "UNDRAFTED" && !result.IsDrafted)
            {
                result.IsDrafted = false;
                continue;
            }
        }

        // Accession: scan from bottom, skip status words, find alphanumeric string >= 10 chars
        // Formats: "317276220260302CRE", "320XR26005765ELP", "SSH2603020023297CST"
        for (int i = lines.Length - 1; i >= 0; i--)
        {
            var line = lines[i].Trim();
            if (line == "DRAFTED" || line == "UNDRAFTED" || line == "Current Study" || line.Length == 0)
                continue;
            if (line.Length >= 10 && System.Text.RegularExpressions.Regex.IsMatch(line, @"^[A-Za-z0-9]+$"))
            {
                result.Accession = line;
            }
            break; // Only check the bottom-most non-status line
        }
    }

    // ═══════ PUBLIC API: EDITOR OPERATIONS ═══════

    /// <summary>
    /// Insert text content into a ProseMirror editor.
    /// editorIndex: 0=transcript, 1=final report.
    /// </summary>
    public bool InsertContent(int editorIndex, string text, bool highlight = false, int[]? mediumConfIndices = null, int[]? lowConfIndices = null)
    {
        if (!IsIframeConnected) return false;
        var escaped = JsonSerializer.Serialize(text); // JSON-escapes the string

        // Build JSON arrays for JS
        var mediumJson = mediumConfIndices is { Length: > 0 }
            ? "[" + string.Join(",", mediumConfIndices) + "]" : "[]";
        var lowJson = lowConfIndices is { Length: > 0 }
            ? "[" + string.Join(",", lowConfIndices) + "]" : "[]";

        string js;
        if (highlight)
        {
            // Insert text then apply CSS Custom Highlights over the inserted range.
            // Uses the CSS Custom Highlight API (CSS.highlights) — purely visual, zero DOM modification.
            // This is the only safe approach that won't trigger Mosaic's change tracking.
            // Three tiers: mt-dictated (all), mt-medium (amber), mt-low (orange).
            js = $@"(() => {{
                const editors = document.querySelectorAll('.ProseMirror');
                if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return 'no_editor';
                const editor = editors[{editorIndex}].editor;
                const posBefore = editor.state.selection.from;
                editor.commands.insertContent({escaped});
                const posAfter = editor.state.selection.from;
                if (posAfter <= posBefore) return 'ok';

                try {{
                    // Get the editor's actual text color for normal words
                    const editorEl = editors[{editorIndex}];
                    const normalColor = getComputedStyle(editorEl).color || 'black';

                    // Ensure CSS rules for our custom highlights exist (always update content)
                    let hlStyle = document.getElementById('mt-dictated-style');
                    if (!hlStyle) {{
                        hlStyle = document.createElement('style');
                        hlStyle.id = 'mt-dictated-style';
                        document.head.appendChild(hlStyle);
                    }}
                    hlStyle.textContent = '::highlight(mt-dictated) {{ background-color: rgba(90, 85, 50, 0.4); }}'
                        + ' ::highlight(mt-normal) {{ color: ' + normalColor + '; }}'
                        + ' ::highlight(mt-medium) {{ color: rgb(200, 170, 80); }}'
                        + ' ::highlight(mt-low) {{ color: rgb(210, 130, 70); }}';

                    // Create a DOM Range covering the inserted text (background highlight)
                    const view = editor.view;
                    const startDOM = view.domAtPos(posBefore);
                    const endDOM = view.domAtPos(posAfter);
                    const range = new Range();
                    range.setStart(startDOM.node, startDOM.offset);
                    range.setEnd(endDOM.node, endDOM.offset);

                    if (!CSS.highlights.has('mt-dictated')) CSS.highlights.set('mt-dictated', new Highlight());
                    CSS.highlights.get('mt-dictated').add(range);

                    // Assign every word an explicit text color (normal/medium/low)
                    const medSet = new Set({mediumJson});
                    const lowSet = new Set({lowJson});
                    const insertedText = editor.state.doc.textBetween(posBefore, posAfter);
                    let wordIndex = 0;
                    let charPos = 0;
                    const len = insertedText.length;
                    while (charPos < len) {{
                        while (charPos < len && insertedText[charPos] === ' ') charPos++;
                        if (charPos >= len) break;
                        const wordStart = charPos;
                        while (charPos < len && insertedText[charPos] !== ' ') charPos++;
                        const hlName = lowSet.has(wordIndex) ? 'mt-low'
                            : medSet.has(wordIndex) ? 'mt-medium' : 'mt-normal';
                        try {{
                            const wStartDOM = view.domAtPos(posBefore + wordStart);
                            const wEndDOM = view.domAtPos(posBefore + charPos);
                            const wRange = new StaticRange({{
                                startContainer: wStartDOM.node, startOffset: wStartDOM.offset,
                                endContainer: wEndDOM.node, endOffset: wEndDOM.offset
                            }});
                            if (!CSS.highlights.has(hlName)) CSS.highlights.set(hlName, new Highlight());
                            CSS.highlights.get(hlName).add(wRange);
                        }} catch(e) {{}}
                        wordIndex++;
                    }}
                    return 'ok_highlight';
                }} catch(e) {{ return 'ok_hl_err:' + e.message; }}
            }})()";
        }
        else
        {
            js = $@"(() => {{
                const editors = document.querySelectorAll('.ProseMirror');
                if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return 'no_editor';
                editors[{editorIndex}].editor.commands.insertContent({escaped});
                return 'ok';
            }})()";
        }

        var result = ExtractResultValue(SendToIframe(js));
        var ok = result != null && result.StartsWith("ok");
        if (ok)
        {
            if (highlight)
                Logger.Trace($"CDP: InsertContent highlight result: {result}");
            ScrollCursorIntoView(editorIndex);
        }
        return ok;
    }

    /// <summary>
    /// Get the character before and after the cursor in a ProseMirror editor,
    /// plus the selected text (if any). Used for smart STT insertions.
    /// </summary>
    public (char before, char after, string selectedText)? GetCursorContext(int editorIndex)
    {
        if (!IsIframeConnected) return null;
        var js = $@"(() => {{
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return null;
            const editor = editors[{editorIndex}].editor;
            const {{ from, to }} = editor.state.selection;
            const doc = editor.state.doc;
            const textBefore = doc.textBetween(Math.max(0, from - 1), from, '\n');
            const textAfter = doc.textBetween(to, Math.min(doc.content.size, to + 1), '\n');
            const sel = from !== to ? doc.textBetween(from, to) : '';
            return JSON.stringify({{ before: textBefore, after: textAfter, sel: sel }});
        }})()";
        var json = ExtractResultValue(SendToIframe(js));
        if (json == null) return null;
        try
        {
            var data = JsonSerializer.Deserialize<JsonElement>(json);
            var before = GetStr(data, "before");
            var after = GetStr(data, "after");
            var sel = GetStr(data, "sel") ?? "";
            char b = string.IsNullOrEmpty(before) ? '\0' : before[0];
            char a = string.IsNullOrEmpty(after) ? '\0' : after[0];
            return (b, a, sel);
        }
        catch { return null; }
    }

    /// <summary>Get the text content of a ProseMirror editor (preserves list numbering).</summary>
    public string? GetEditorText(int editorIndex)
    {
        if (!IsIframeConnected) return null;
        var js = $@"(() => {{
            function walk(node) {{
                let t = '';
                for (const c of node.childNodes) {{
                    if (c.nodeType === 3) {{ t += c.textContent; continue; }}
                    const tag = c.tagName;
                    if (tag === 'BR') {{ t += '\n'; continue; }}
                    if (tag === 'OL') {{
                        let n = parseInt(c.getAttribute('start')) || 1;
                        for (const li of c.children) {{
                            if (li.tagName === 'LI') {{ t += n + '. ' + walk(li).trim() + '\n'; n++; }}
                        }}
                        continue;
                    }}
                    if (tag === 'P' || tag === 'DIV' || (tag && tag.match(/^H[1-6]$/))) {{
                        t += walk(c) + '\n'; continue;
                    }}
                    t += walk(c);
                }}
                return t;
            }}
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length <= {editorIndex}) return null;
            return walk(editors[{editorIndex}]);
        }})()";
        return ExtractResultValue(SendToIframe(js));
    }

    /// <summary>
    /// Get the focused editor index. First tries live DOM detection, then falls back to
    /// the last-known focused editor tracked during scrape (survives window backgrounding).
    /// Returns -1 only if never detected.
    /// </summary>
    public int GetFocusedEditorIndex()
    {
        if (!IsIframeConnected) return _lastKnownFocusedEditor;
        var js = @"(() => {
            const editors = document.querySelectorAll('.ProseMirror');
            for (let i = 0; i < editors.length; i++) {
                if (editors[i].classList.contains('ProseMirror-focused')) return i;
            }
            for (let i = 0; i < editors.length; i++) {
                if (editors[i].contains(document.activeElement) || editors[i] === document.activeElement) return i;
            }
            return -1;
        })()";
        var result = ExtractResultValue(SendToIframe(js));
        if (result != null && int.TryParse(result, out int idx) && idx >= 0)
        {
            _lastKnownFocusedEditor = idx;
            return idx;
        }
        // Live detection failed (window backgrounded) — use last-known from scrape
        return _lastKnownFocusedEditor;
    }

    /// <summary>Focus a ProseMirror editor at the end.</summary>
    public bool FocusEditor(int editorIndex)
    {
        if (!IsIframeConnected) return false;
        var js = $@"(() => {{
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return 'no_editor';
            editors[{editorIndex}].editor.commands.focus('end');
            return 'ok';
        }})()";
        return ExtractResultValue(SendToIframe(js)) == "ok";
    }

    /// <summary>Select the IMPRESSION section content in editor 1 (final report).</summary>
    public bool SelectImpressionContent()
    {
        if (!IsIframeConnected) return false;
        var js = @"(() => {
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length < 2 || !editors[1].editor) return 'no_editor';
            const editor = editors[1].editor;
            const doc = editor.state.doc;
            // Find text node containing 'IMPRESSION', resolve to its parent paragraph,
            // then select from end of that paragraph to end of doc
            let selStart = -1;
            doc.descendants((node, pos) => {
                if (selStart >= 0) return false;
                if (node.isText && node.text && node.text.includes('IMPRESSION')) {
                    // Resolve position to find parent block boundary
                    const $pos = doc.resolve(pos);
                    // $pos.end() = end of innermost block, +1 = start of next block
                    selStart = $pos.end($pos.depth) + 1;
                    return false;
                }
            });
            if (selStart < 0) return 'not_found';
            const docEnd = doc.content.size - 1;
            if (selStart >= docEnd) return 'empty';
            editor.commands.setTextSelection({ from: selStart, to: docEnd });
            return 'ok';
        })()";
        return ExtractResultValue(SendToIframe(js)) == "ok";
    }

    // ═══════ EXPERIMENTAL: EDITOR HIGHLIGHTING ═══════

    /// <summary>
    /// [Test] Highlight all occurrences of a word in the editor with a custom color.
    /// Returns count of occurrences found.
    /// </summary>
    public int HighlightWord(int editorIndex, string word, string color = "#FF6B35")
    {
        if (!IsIframeConnected) return -1;
        var escapedWord = word.Replace("'", "\\'").Replace("\\", "\\\\");
        var js = $@"(() => {{
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return -1;
            const editor = editors[{editorIndex}].editor;
            const doc = editor.state.doc;
            const word = '{escapedWord}'.toLowerCase();
            let found = [];
            doc.descendants((node, pos) => {{
                if (node.isText) {{
                    let idx = node.text.toLowerCase().indexOf(word);
                    while (idx !== -1) {{
                        found.push({{ from: pos + idx, to: pos + idx + word.length }});
                        idx = node.text.toLowerCase().indexOf(word, idx + 1);
                    }}
                }}
            }});
            if (found.length === 0) return 0;
            // Apply highlight to each occurrence (reverse order to preserve positions)
            for (let i = found.length - 1; i >= 0; i--) {{
                editor.chain()
                    .setTextSelection(found[i])
                    .setHighlight({{ color: '{color}' }})
                    .run();
            }}
            editor.commands.setTextSelection(doc.content.size - 1);
            return found.length;
        }})()";
        var result = ExtractResultValue(SendToIframe(js));
        if (result != null && int.TryParse(result, out var count))
        {
            Logger.Trace($"CDP HighlightWord: '{word}' color={color} → {count} occurrences");
            return count;
        }
        return -1;
    }

    /// <summary>
    /// [Test] Remove all highlights from the editor (restores to unhighlighted state).
    /// </summary>
    public bool ClearAllHighlights(int editorIndex)
    {
        if (!IsIframeConnected) return false;
        var js = $@"(() => {{
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return false;
            const editor = editors[{editorIndex}].editor;
            editor.chain().selectAll().unsetHighlight().run();
            editor.commands.setTextSelection(editor.state.doc.content.size - 1);
            return true;
        }})()";
        var result = ExtractResultValue(SendToIframe(js));
        return result == "true";
    }

    /// <summary>
    /// Capture the ProseMirror doc structure (heading levels, section/orderedList wrapping)
    /// so we can reconstruct it when writing LLM output back.
    /// Returns JSON: { headings: { "FINDINGS": 4, "IMPRESSION": 4, ... }, hasImpressionSection: true }
    /// </summary>
    public string? GetDocStructure(int editorIndex)
    {
        if (!IsIframeConnected) return null;
        var js = $@"(() => {{
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return null;
            const doc = editors[{editorIndex}].editor.getJSON();
            if (!doc || !doc.content) return null;
            const headings = {{}};
            let hasImpressionSection = false;
            for (const node of doc.content) {{
                if (node.type === 'heading') {{
                    let text = '';
                    if (node.content) node.content.forEach(c => {{ if (c.text) text += c.text; }});
                    text = text.replace(/:$/, '').trim().toUpperCase();
                    if (text) headings[text] = node.attrs?.level || 4;
                }}
                if (node.type === 'section') {{
                    hasImpressionSection = true;
                    if (node.content) {{
                        for (const child of node.content) {{
                            if (child.type === 'heading') {{
                                let text = '';
                                if (child.content) child.content.forEach(c => {{ if (c.text) text += c.text; }});
                                text = text.replace(/:$/, '').trim().toUpperCase();
                                if (text) headings[text] = child.attrs?.level || 4;
                            }}
                        }}
                    }}
                }}
            }}
            return JSON.stringify({{ headings, hasImpressionSection }});
        }})()";
        return ExtractResultValue(SendToIframe(js));
    }

    /// <summary>
    /// Replace editor content preserving ProseMirror heading/section/orderedList structure.
    /// docStructureJson: output from GetDocStructure (heading levels, impression section flag).
    /// </summary>
    public bool ReplaceEditorContent(int editorIndex, string text, string? docStructureJson = null)
    {
        if (!IsIframeConnected) return false;
        var escapedText = JsonSerializer.Serialize(text);
        var escapedStructure = docStructureJson != null ? JsonSerializer.Serialize(docStructureJson) : "null";
        var js = $@"(() => {{
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length <= {editorIndex} || !editors[{editorIndex}].editor) return 'no_editor';
            const editor = editors[{editorIndex}].editor;
            const text = {escapedText};
            const structStr = {escapedStructure};
            const struct_ = structStr ? JSON.parse(structStr) : null;
            const headingLevels = struct_?.headings || {{}};
            const hasImpSection = struct_?.hasImpressionSection || false;

            // Build set of ALL known headings from the original doc structure
            const knownHeadings = new Set(Object.keys(headingLevels));

            const lines = text.split('\n');
            const content = [];
            let i = 0;

            function isHeading(line) {{
                const trimmed = line.trim().replace(/:$/, '').trim().toUpperCase();
                return knownHeadings.has(trimmed);
            }}

            function getHeadingLevel(name) {{
                return headingLevels[name.toUpperCase()] || 4;
            }}

            function makeHeading(name, level) {{
                return {{ type: 'heading', attrs: {{ level: level }}, content: [{{ type: 'text', text: name + ':' }}] }};
            }}

            function makeParagraph(line) {{
                if (!line || !line.trim()) return {{ type: 'paragraph' }};
                return {{ type: 'paragraph', content: [{{ type: 'text', text: line }}] }};
            }}

            // Collect consecutive body lines between headings into single paragraphs.
            // In radiology reports, text under a subsection heading is one paragraph.
            // Blank lines create paragraph breaks.
            function collectBody() {{
                const parts = [];
                while (i < lines.length) {{
                    const t = lines[i].trim();
                    if (!t) {{ i++; if (parts.length > 0) break; continue; }} // blank line = paragraph break
                    if (isHeading(t)) break; // next heading — stop collecting
                    parts.push(t);
                    i++;
                }}
                return parts.length > 0 ? parts.join(' ') : null;
            }}

            while (i < lines.length) {{
                const line = lines[i];
                const trimmed = line.trim();

                // Skip blank lines at top level
                if (!trimmed) {{ i++; continue; }}

                if (isHeading(trimmed)) {{
                    const headerName = trimmed.replace(/:$/, '').trim();
                    const upperName = headerName.toUpperCase();
                    const level = getHeadingLevel(headerName);

                    // Add empty paragraph spacer before headings (not before the first one)
                    if (content.length > 0) {{
                        content.push({{ type: 'paragraph' }});
                    }}

                    // IMPRESSION with section wrapper + orderedList
                    if (upperName === 'IMPRESSION' && hasImpSection) {{
                        i++;
                        const impItems = [];
                        while (i < lines.length && !isHeading(lines[i])) {{
                            const impLine = lines[i].trim();
                            if (impLine) {{
                                const stripped = impLine.replace(/^\d+\.\s*/, '');
                                if (stripped) {{
                                    impItems.push({{
                                        type: 'listItem',
                                        content: [{{ type: 'paragraph', content: [{{ type: 'text', text: stripped }}] }}]
                                    }});
                                }}
                            }}
                            i++;
                        }}
                        const sectionContent = [makeHeading(headerName, level)];
                        if (impItems.length > 0) {{
                            sectionContent.push({{
                                type: 'orderedList',
                                attrs: {{ start: 1 }},
                                content: impItems
                            }});
                        }}
                        content.push({{ type: 'section', content: sectionContent }});
                        continue;
                    }}

                    // Regular heading (top-level or subsection)
                    content.push(makeHeading(headerName, level));
                    i++;

                    // Collect body text after heading into single paragraph(s)
                    let body;
                    while ((body = collectBody()) !== null) {{
                        content.push(makeParagraph(body));
                    }}
                    continue;
                }}

                // Body text not under a heading — collect into paragraph
                const body = collectBody();
                if (body) content.push(makeParagraph(body));
            }}

            if (content.length === 0) content.push({{ type: 'paragraph' }});
            editor.commands.setContent({{ type: 'doc', content: content }});

            // Nudge Mosaic's change detection — setContent() bypasses the normal editing
            // pipeline. Insert a zero-width space then immediately delete it in a separate
            // transaction, which triggers ProseMirror's plugin/listener machinery without
            // leaving any trace. Use addToHistory:false so it doesn't pollute undo stack.
            const view = editor.view;
            const endPos = editor.state.doc.content.size - 1;
            const tr1 = view.state.tr.insertText('\u200B', endPos);
            tr1.setMeta('addToHistory', false);
            view.dispatch(tr1);
            const tr2 = view.state.tr.delete(endPos, endPos + 1);
            tr2.setMeta('addToHistory', false);
            view.dispatch(tr2);

            return 'ok';
        }})()";
        var result = ExtractResultValue(SendToIframe(js));
        if (result != "ok")
            Logger.Trace($"CDP: ReplaceEditorContent({editorIndex}) failed: {result}");
        return result == "ok";
    }

    /// <summary>
    /// Flash matching final-report text for active alert terms.
    /// Installs a persistent JS watchdog interval that re-applies ProseMirror highlight
    /// marks whenever Mosaic's internal sync strips them (~every 2s).
    /// CSS @keyframes animation on the resulting mark elements creates the visual flash.
    /// </summary>
    public bool UpdateAlertTextFlashing(IReadOnlyList<string>? genderTerms, IReadOnlyList<string>? fimTerms)
    {
        if (!IsIframeConnected) return false;

        var gender = (genderTerms ?? Array.Empty<string>())
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var fim = (fimTerms ?? Array.Empty<string>())
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var genderJson = JsonSerializer.Serialize(gender);
        var fimJson = JsonSerializer.Serialize(fim);

        // Store terms on window.__mtFlash and install a watchdog setInterval that
        // re-applies marks every 500ms when Mosaic's document sync strips them.
        var js = $@"(() => {{
            const GC = '#DC0000', FC = '#FF8C00', SID = 'mt-alert-flash-css';

            // Find the report editor (highest scoring ProseMirror instance)
            function findEditor() {{
                const eds = Array.from(document.querySelectorAll('.ProseMirror'));
                if (eds.length === 0) return null;
                function sc(ed) {{
                    const t = (ed.innerText || '').toUpperCase();
                    let s = 0;
                    if (ed.isConnected && ed.getClientRects().length > 0) s += 3;
                    if (t.includes('IMPRESSION')) s += 5;
                    if (t.includes('FINDINGS')) s += 4;
                    if (t.includes('EXAM')) s += 2;
                    if (t.length > 200) s += 1;
                    return s;
                }}
                let best = eds[0], bestSc = sc(eds[0]);
                for (let i = 1; i < eds.length; i++) {{
                    const s = sc(eds[i]);
                    if (s > bestSc) {{ bestSc = s; best = eds[i]; }}
                }}
                return best.editor || null;
            }}

            // Apply highlight marks for given terms with given color
            function applyMarks(tip, hl, terms, color, tr) {{
                let c = 0;
                const doc = tr.doc;
                terms.forEach(t => {{
                    const re = new RegExp('\\b' + t.replace(/[.*+?^${{}}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+') + '\\b', 'gi');
                    doc.descendants((n, p) => {{
                        if (!n.isText) return;
                        re.lastIndex = 0;
                        let m;
                        while ((m = re.exec(n.text)) !== null) {{
                            tr.addMark(p + m.index, p + m.index + m[0].length, hl.create({{ color }}));
                            c++;
                        }}
                    }});
                }});
                return c;
            }}

            // Check if our marks exist in the document
            function hasOurMarks(tip, hl) {{
                let found = false;
                tip.state.doc.descendants(n => {{
                    if (found || !n.isText) return;
                    n.marks.forEach(mk => {{
                        if (mk.type === hl && (mk.attrs.color === GC || mk.attrs.color === FC))
                            found = true;
                    }});
                }});
                return found;
            }}

            // Core: remove old marks then apply new ones
            function refreshMarks() {{
                const f = window.__mtFlash;
                if (!f || !f.active) return;
                const tip = findEditor();
                if (!tip) return;
                const hl = tip.schema.marks.highlight;
                if (!hl) return;

                // Skip if our marks are still present
                if (hasOurMarks(tip, hl)) return;

                const tr = tip.state.tr;
                tr.setMeta('addToHistory', false);
                applyMarks(tip, hl, f.gT, GC, tr);
                applyMarks(tip, hl, f.fT, FC, tr);
                if (tr.steps.length > 0) tip.view.dispatch(tr);
            }}

            // Update stored terms
            if (!window.__mtFlash) window.__mtFlash = {{}};
            const f = window.__mtFlash;
            f.gT = {genderJson};
            f.fT = {fimJson};
            f.active = true;

            // Ensure CSS flash animation exists
            let sEl = document.getElementById(SID);
            if (!sEl) {{ sEl = document.createElement('style'); sEl.id = SID; document.head.appendChild(sEl); }}
            sEl.textContent =
                '@keyframes mt-fg{{0%,50%{{background-color:' + GC + '}}51%,100%{{background-color:transparent}}}}' +
                '@keyframes mt-ff{{0%,50%{{background-color:' + FC + '}}51%,100%{{background-color:transparent}}}}' +
                'mark[data-color=""' + GC + '""],mark[data-color=""' + GC + '""] *{{animation:mt-fg 1s infinite!important;color:white!important;-webkit-text-fill-color:white!important}}' +
                'mark[data-color=""' + FC + '""],mark[data-color=""' + FC + '""] *{{animation:mt-ff 1s infinite!important;color:white!important;-webkit-text-fill-color:white!important}}';

            // Install watchdog interval (or keep existing one)
            if (f.iv) clearInterval(f.iv);
            f.iv = setInterval(refreshMarks, 500);

            // Apply immediately on this call
            const tip = findEditor();
            if (!tip) return 'no_tiptap';
            const hl = tip.schema.marks.highlight;
            if (!hl) return 'no_highlight';
            const tr = tip.state.tr;
            tr.setMeta('addToHistory', false);

            // Remove stale marks first
            tip.state.doc.descendants((node, pos) => {{
                if (!node.isText) return;
                node.marks.forEach(mk => {{
                    if (mk.type === hl && (mk.attrs.color === GC || mk.attrs.color === FC))
                        tr.removeMark(pos, pos + node.nodeSize, mk);
                }});
            }});

            const gC = applyMarks(tip, hl, f.gT, GC, tr);
            const fC = applyMarks(tip, hl, f.fT, FC, tr);
            if (tr.steps.length > 0) tip.view.dispatch(tr);

            return JSON.stringify({{ g: gC, f: fC, watchdog: true }});
        }})()";

        var result = ExtractResultValue(SendToIframe(js));
        if (result == null) return false;
        Logger.Trace($"CDP: Alert text flashing update: {result}");
        return !result.StartsWith("no_", StringComparison.Ordinal);
    }

    /// <summary>
    /// Remove alert-text flashing highlights and stop the watchdog interval.
    /// </summary>
    public void ClearAlertTextFlashing()
    {
        if (!IsIframeConnected) return;
        var js = @"(() => {
            const GC = '#DC0000', FC = '#FF8C00';
            // Stop watchdog
            if (window.__mtFlash) {
                if (window.__mtFlash.iv) clearInterval(window.__mtFlash.iv);
                window.__mtFlash = null;
            }
            // Remove marks from all editors
            document.querySelectorAll('.ProseMirror').forEach(el => {
                const tip = el.editor;
                if (!tip) return;
                const hl = tip.schema.marks.highlight;
                if (!hl) return;
                const doc = tip.state.doc;
                const tr = tip.state.tr;
                tr.setMeta('addToHistory', false);
                doc.descendants((node, pos) => {
                    if (!node.isText) return;
                    node.marks.forEach(mk => {
                        if (mk.type === hl && (mk.attrs.color === GC || mk.attrs.color === FC))
                            tr.removeMark(pos, pos + node.nodeSize, mk);
                    });
                });
                if (tr.steps.length > 0) tip.view.dispatch(tr);
            });
            const style = document.getElementById('mt-alert-flash-css');
            if (style) style.remove();
            return 'ok';
        })()";
        try { SendToIframe(js); } catch { }
    }

    // ═══════ PUBLIC API: STRUCTURED REPORT ═══════

    // JS payload that walks ProseMirror JSON tree and returns structured sections.
    // Highlighted text = user-dictated content. Non-highlighted = template boilerplate.
    // IMPRESSION items are always treated as dictated regardless of highlight state.
    // Runs entirely in browser — only compact result crosses CDP.
    private const string JS_GET_STRUCTURED_REPORT = @"(() => {
        const editors = document.querySelectorAll('.ProseMirror');
        if (editors.length < 2 || !editors[1].editor) return null;
        const doc = editors[1].editor.getJSON();
        if (!doc || !doc.content) return null;

        const sections = [];
        let current = null;
        let inFindings = false;
        let currentSub = null;

        function textFromNode(node, dictOnly, tmplOnly) {
            if (!node) return '';
            if (node.type === 'text') {
                const hasHL = node.marks && node.marks.some(m => m.type === 'highlight');
                // Highlighted = dictated, non-highlighted = template
                if (dictOnly && !hasHL) return '';
                if (tmplOnly && hasHL) return '';
                return node.text || '';
            }
            if (!node.content) return '';
            return node.content.map(c => textFromNode(c, dictOnly, tmplOnly)).join('');
        }

        function allText(node) { return textFromNode(node, false, false); }
        function dictText(node) { return textFromNode(node, true, false); }
        function tmplText(node) { return textFromNode(node, false, true); }

        function collectContent(node) {
            const full = allText(node).trim();
            const dict = dictText(node).trim();
            const tmpl = tmplText(node).trim();
            return { full, dict, tmpl };
        }

        function pushContent(target, node) {
            const c = collectContent(node);
            if (c.full) {
                target.fullText += (target.fullText ? '\n' : '') + c.full;
                // If ANY part of a paragraph is highlighted, the whole paragraph is dictated
                if (c.dict) target.dictatedText += (target.dictatedText ? '\n' : '') + c.full;
                if (c.tmpl) target.templateText += (target.templateText ? '\n' : '') + c.tmpl;
            }
        }

        for (const node of doc.content) {
            if (node.type === 'heading') {
                const name = allText(node).replace(/:$/, '').trim();
                if (!name) continue;

                // Is this a top-level section or a FINDINGS subsection?
                const isTopLevel = ['EXAM','TECHNIQUE','COMPARISON','CLINICAL HISTORY',
                    'FINDINGS','IMPRESSION','INDICATION','PROCEDURE','CONCLUSION',
                    'RECOMMENDATION','SIGNATURE','ADDENDUM'].includes(name.toUpperCase());

                if (isTopLevel || !inFindings) {
                    currentSub = null;
                    current = { name: name, fullText: '', dictatedText: '', templateText: '', items: [], subsections: [] };
                    sections.push(current);
                    inFindings = (name.toUpperCase() === 'FINDINGS');
                } else {
                    // Subsection under FINDINGS
                    currentSub = { name: name, fullText: '', dictatedText: '', templateText: '', items: [], subsections: [] };
                    if (current) current.subsections.push(currentSub);
                }
                continue;
            }

            if (node.type === 'section') {
                // IMPRESSION is wrapped in a <section> node
                if (node.content) {
                    let impSection = null;
                    for (const child of node.content) {
                        if (child.type === 'heading') {
                            const name = allText(child).replace(/:$/, '').trim();
                            impSection = { name: name, fullText: '', dictatedText: '', templateText: '', items: [], subsections: [] };
                            sections.push(impSection);
                            current = impSection;
                            currentSub = null;
                            inFindings = false;
                        } else if (child.type === 'orderedList' && child.content && impSection) {
                            let n = (child.attrs && child.attrs.start) || 1;
                            for (const li of child.content) {
                                if (li.type === 'listItem') {
                                    const itemText = allText(li).trim();
                                    if (itemText) {
                                        impSection.items.push(itemText);
                                        const numbered = n + '. ' + itemText;
                                        impSection.fullText += (impSection.fullText ? '\n' : '') + numbered;
                                        // Impression items are always dictated
                                        impSection.dictatedText += (impSection.dictatedText ? '\n' : '') + numbered;
                                    }
                                    n++;
                                }
                            }
                        } else if (impSection) {
                            pushContent(impSection, child);
                        }
                    }
                }
                continue;
            }

            // Regular content node (draggableItem, draggableInline, paragraph)
            if (!current) continue;
            const target = currentSub || current;
            pushContent(target, node);
        }

        return JSON.stringify(sections);
    })()";

    /// <summary>
    /// Get the final report as a structured object with sections, dictated/template text separation,
    /// and IMPRESSION items. Uses ProseMirror's getJSON() for reliable parsing instead of regex.
    /// Optional enhancement — existing regex-based parsing remains as fallback.
    /// </summary>
    public StructuredReport? GetStructuredReport()
    {
        if (!IsIframeConnected) return null;

        var json = ExtractResultValue(SendToIframe(JS_GET_STRUCTURED_REPORT));
        if (json == null) return null;

        try
        {
            var sections = JsonSerializer.Deserialize<List<ReportSection>>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            if (sections == null) return null;

            var report = new StructuredReport { Sections = sections };
            Logger.Trace($"CDP: GetStructuredReport — {sections.Count} sections: {string.Join(", ", sections.Select(s => s.Name))}");
            return report;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: GetStructuredReport parse error: {ex.Message}");
            return null;
        }
    }

    // ═══════ PUBLIC API: BUTTON CLICKS ═══════

    /// <summary>Click the Process Report button in the iframe.</summary>
    public bool ClickProcessReport()
    {
        if (!IsIframeConnected) return false;
        var js = @"(() => {
            const btns = document.querySelectorAll('button');
            for (const b of btns) {
                if (b.innerText?.trim() === 'Process Report' || b.getAttribute('aria-label') === 'Process Report') {
                    b.click(); return 'ok';
                }
            }
            return 'not_found';
        })()";
        return ExtractResultValue(SendToIframe(js)) == "ok";
    }

    /// <summary>Click the Sign Report button in the iframe.</summary>
    public bool ClickSignReport()
    {
        if (!IsIframeConnected) return false;
        var js = @"(() => {
            const btns = document.querySelectorAll('button');
            for (const b of btns) {
                if (b.innerText?.trim() === 'Sign Report' || b.getAttribute('aria-label') === 'Sign Report') {
                    b.click(); return 'ok';
                }
            }
            return 'not_found';
        })()";
        return ExtractResultValue(SendToIframe(js)) == "ok";
    }

    /// <summary>Click the Create Impression button in the iframe.</summary>
    public bool ClickCreateImpression()
    {
        if (!IsIframeConnected) return false;
        var js = @"(() => {
            const btns = document.querySelectorAll('button');
            for (const b of btns) {
                const text = b.innerText?.trim() || '';
                if (text === 'Create Impression' || text.includes('Impression')) {
                    b.click(); return 'ok';
                }
            }
            return 'not_found';
        })()";
        return ExtractResultValue(SendToIframe(js)) == "ok";
    }

    /// <summary>Click the Discard Study button in the iframe.</summary>
    public bool ClickDiscardStudy()
    {
        if (!IsIframeConnected) return false;
        var js = @"(() => {
            const btns = document.querySelectorAll('button');
            for (const b of btns) {
                const text = b.innerText?.trim() || '';
                const label = b.getAttribute('aria-label') || '';
                if (text === 'Discard' || text === 'Discard Study' || label.includes('Discard')) {
                    b.click(); return 'ok';
                }
            }
            return 'not_found';
        })()";
        return ExtractResultValue(SendToIframe(js)) == "ok";
    }

    // ═══════ PUBLIC API: SCROLL FIX ═══════

    /// <summary>
    /// Inject CSS into the iframe to make the three columns independently scrollable.
    /// Always re-discovers DOM and injects fresh CSS (no stale persistence).
    /// Checks if style tag still exists in DOM before skipping — handles study changes that refresh DOM.
    /// </summary>
    public bool InjectScrollFix()
    {
        if (_scrollFixActive)
        {
            // Verify the style tag still exists (DOM refreshes on study change)
            try
            {
                var check = ExtractResultValue(SendToIframe(
                    "!!document.getElementById('mt-scroll-fix')"));
                if (check == "true") return true;
                // Style tag gone — DOM was refreshed, re-inject
                _scrollFixActive = false;
                Logger.Trace("CDP: Scroll fix style tag gone — re-injecting");
            }
            catch { _scrollFixActive = false; }
        }
        if (!IsIframeConnected) return false;

        // Format template with current column ratio and drag handle hiding flag
        var js = string.Format(JS_INJECT_SCROLL_FIX_TEMPLATE,
            _columnRatio.ToString(System.Globalization.CultureInfo.InvariantCulture),
            _hideDragHandles ? "true" : "false");
        var resultJson = ExtractResultValue(SendToIframe(js));
        if (resultJson == null) return false;

        try
        {
            var data = JsonSerializer.Deserialize<JsonElement>(resultJson);
            if (data.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                _scrollFixActive = true;
                var layout = GetStr(data, "layout") ?? "?";
                var cols = data.TryGetProperty("columns", out var c) ? c.ToString() : "?";
                var top = data.TryGetProperty("topPx", out var t) ? t.ToString() : "?";
                var areas = data.TryGetProperty("editorAreas", out var ea) ? ea.ToString() : "?";
                var wrappers = data.TryGetProperty("wrappers", out var w) ? w.ToString() : "";
                var containerW = data.TryGetProperty("containerW", out var cw) ? cw.ToString() : "?";
                var windowW = data.TryGetProperty("windowW", out var ww) ? ww.ToString() : "?";
                var colWidths = data.TryGetProperty("colWidths", out var cwArr) ? cwArr.ToString() : "";
                Logger.Trace($"CDP: Scroll fix injected — {layout}, {cols} columns, {areas} editor areas, topPx={top}, ratio={_columnRatio:F3}, containerW={containerW}, windowW={windowW}, colWidths={colWidths}, wrappers={wrappers}");
                return true;
            }
            // Not ready yet (e.g., editors not loaded) — will retry next tick
            var error = GetStr(data, "error") ?? "unknown";
            Logger.Trace($"CDP: Scroll fix not ready: {error}");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: Scroll fix error: {ex.Message}");
            return false;
        }
    }

    /// <summary>Dump DOM structure around editors for debugging scroll/resize issues.</summary>
    public string? DumpScrollDiagnostic()
    {
        if (!IsIframeConnected) return null;
        var js = @"(() => {
            const editors = document.querySelectorAll('.ProseMirror');
            if (editors.length < 2) return JSON.stringify({error:'need_2_editors'});
            const result = [];
            for (let ei = 0; ei < Math.min(editors.length, 2); ei++) {
                const chain = [];
                let el = editors[ei];
                while (el && el !== document.body) {
                    const cs = getComputedStyle(el);
                    chain.push({
                        tag: el.tagName,
                        cls: (el.className||'').substring(0,80),
                        id: el.id||'',
                        w: el.offsetWidth, h: el.offsetHeight,
                        ov: cs.overflow, ovx: cs.overflowX, ovy: cs.overflowY,
                        disp: cs.display, flexDir: cs.flexDirection,
                        flexBasis: cs.flexBasis, flexGrow: cs.flexGrow, flexShrink: cs.flexShrink,
                        mt: Array.from(el.attributes).filter(a=>a.name.startsWith('data-mt')).map(a=>a.name).join(',')
                    });
                    el = el.parentElement;
                }
                result.push({editor: ei, chain: chain});
            }
            // Also dump flex container children widths
            const cols = document.querySelector('[data-mt-cols]');
            if (cols) {
                const kids = Array.from(cols.children).map(c => ({
                    tag: c.tagName, cls: (c.className||'').substring(0,50),
                    w: c.offsetWidth, h: c.offsetHeight,
                    flexBasis: getComputedStyle(c).flexBasis,
                    flexGrow: getComputedStyle(c).flexGrow,
                    flexShrink: getComputedStyle(c).flexShrink,
                    mt: Array.from(c.attributes).filter(a=>a.name.startsWith('data-mt')).map(a=>a.name).join(',')
                }));
                result.push({colContainer: true, w: cols.offsetWidth, kids: kids});
            }
            return JSON.stringify(result);
        })()";
        return ExtractResultValue(SendToIframe(js));
    }

    /// <summary>Dump children of tagged elements for debugging vertical layout issues.</summary>
    public string? DumpVerticalLayoutDiag()
    {
        if (!IsIframeConnected) return null;
        // Walk from each ProseMirror UP to its wrapper, dumping each ancestor's layout info
        // plus its siblings — reveals what creates the internal two-column layout
        var js = @"(() => {
            const result = [];
            const editors = document.querySelectorAll('.ProseMirror');
            editors.forEach((editor, ei) => {
                const chain = [];
                let el = editor;
                const wrapper = el.closest('[data-mt-editor-wrapper]');
                while (el && el !== wrapper) {
                    const parent = el.parentElement;
                    if (!parent) break;
                    const ps = getComputedStyle(parent);
                    const siblings = Array.from(parent.children).map(s => {
                        const ss = getComputedStyle(s);
                        return {
                            tag: s.tagName,
                            cls: (s.className||'').substring(0,60),
                            w: s.offsetWidth, h: s.offsetHeight,
                            disp: ss.display, flexDir: ss.flexDirection,
                            maxW: ss.maxWidth, gridCol: ss.gridTemplateColumns?.substring(0,40) || '',
                            mt: Array.from(s.attributes).filter(a=>a.name.startsWith('data-mt')).map(a=>a.name).join(','),
                            isPath: s === el
                        };
                    });
                    chain.push({
                        parentTag: parent.tagName,
                        parentCls: (parent.className||'').substring(0,80),
                        parentW: parent.offsetWidth, parentH: parent.offsetHeight,
                        disp: ps.display, flexDir: ps.flexDirection,
                        gridCols: ps.gridTemplateColumns?.substring(0,60) || '',
                        maxW: ps.maxWidth,
                        mt: Array.from(parent.attributes).filter(a=>a.name.startsWith('data-mt')).map(a=>a.name).join(','),
                        childCount: siblings.length,
                        children: siblings
                    });
                    el = parent;
                }
                result.push({editor: ei, levels: chain.length, chain});
            });
            return JSON.stringify(result);
        })()";
        return ExtractResultValue(SendToIframe(js));
    }

    /// <summary>Find interactive/icon elements inside ProseMirror editors for diagnostics.</summary>
    public string? DumpEditorControls()
    {
        if (!IsIframeConnected) return null;
        var js = @"(() => {
            const pm = document.querySelectorAll('.ProseMirror');
            const results = [];
            pm.forEach((editor, ei) => {
                // Find SVGs, buttons, clickable icons, draggable elements, small interactive elements
                const candidates = editor.querySelectorAll('button, [role=""button""], svg, [draggable], [data-drag-handle], [data-delete-handle], [contenteditable=""false""]');
                candidates.forEach(el => {
                    const cs = getComputedStyle(el);
                    if (cs.display === 'none') return; // skip already hidden
                    const attrs = Array.from(el.attributes).map(a => a.name + '=' + a.value.substring(0,30)).join(', ');
                    const parentTag = el.parentElement ? el.parentElement.tagName : '';
                    const parentCls = el.parentElement ? (el.parentElement.className||'').substring(0,60) : '';
                    const parentAttrs = el.parentElement ? Array.from(el.parentElement.attributes).filter(a=>a.name!=='class'&&a.name!=='style').map(a => a.name + '=' + a.value.substring(0,30)).join(', ') : '';
                    results.push({
                        editor: ei,
                        tag: el.tagName,
                        cls: (el.className&&el.className.substring ? el.className.substring(0,80) : ''),
                        w: el.offsetWidth, h: el.offsetHeight,
                        attrs: attrs.substring(0,200),
                        parentTag, parentCls, parentAttrs,
                        disp: cs.display, pos: cs.position
                    });
                });
            });
            return JSON.stringify(results);
        })()";
        return ExtractResultValue(SendToIframe(js));
    }

    /// <summary>Read back the column ratio from the browser (after user drag) and persist to config.</summary>
    public double? ReadColumnRatio()
    {
        if (!IsIframeConnected) return null;
        var result = ExtractResultValue(SendToIframe("window.__mtColumnRatio || null"));
        if (result != null && double.TryParse(result, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var ratio))
        {
            _columnRatio = Math.Clamp(ratio, 0.15, 0.75);
            return _columnRatio;
        }
        return null;
    }

    /// <summary>Remove the injected scroll fix CSS from the iframe.</summary>
    public void RemoveScrollFix()
    {
        _scrollFixActive = false;
        if (!IsIframeConnected) return;
        try { SendToIframe(JS_REMOVE_SCROLL_FIX); }
        catch { }
    }

    // ═══════ PUBLIC API: MOSAIC MACROS ═══════

    /// <summary>
    /// Fetch Mosaic macro definitions from the RadPair API via the iframe's auth context.
    /// Called once after iframe connects. Results cached in MosaicMacros dictionary.
    /// </summary>
    public void FetchMosaicMacros()
    {
        if (!IsIframeConnected || _mosaicMacrosFetched) return;
        _mosaicMacrosFetched = true;

        // Execute fetch() inside the iframe — piggybacks on the page's existing auth cookies/headers
        var js = @"(async () => {
            try {
                const resp = await fetch('https://api-rp.radpair.com/macros/available', { credentials: 'include' });
                if (!resp.ok) return JSON.stringify({ error: resp.status });
                const data = await resp.json();
                return JSON.stringify(data);
            } catch(e) { return JSON.stringify({ error: e.message }); }
        })()";

        // Runtime.evaluate with awaitPromise for async fetch
        var result = SendToIframeAsync(js);
        if (result == null)
        {
            Logger.Trace("CDP: FetchMosaicMacros — no response");
            return;
        }

        try
        {
            var root = JsonSerializer.Deserialize<JsonElement>(result);

            // Check for error
            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("error", out var err))
            {
                Logger.Trace($"CDP: FetchMosaicMacros — API error: {err}");
                _mosaicMacrosFetched = false; // Allow retry
                return;
            }

            // Parse macro array — we don't know the exact schema yet, so log it and try common patterns
            var macros = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (root.ValueKind == JsonValueKind.Array)
            {
                ParseMacroArray(root, macros);
            }
            else if (root.ValueKind == JsonValueKind.Object)
            {
                // Might be { results: [...] } or { macros: [...] } wrapper
                foreach (var prop in root.EnumerateObject())
                {
                    if (prop.Value.ValueKind == JsonValueKind.Array)
                    {
                        ParseMacroArray(prop.Value, macros);
                        break;
                    }
                }
            }

            MosaicMacros = macros;
            Logger.Trace($"CDP: FetchMosaicMacros — loaded {macros.Count} Mosaic macros");
            if (macros.Count > 0)
            {
                // Log first few names for debugging
                var sample = string.Join(", ", macros.Keys.Take(5));
                Logger.Trace($"CDP: Mosaic macro names (sample): {sample}");
            }
            else
            {
                // Log raw response for schema discovery
                var preview = result.Length > 500 ? result[..500] + "..." : result;
                Logger.Trace($"CDP: Mosaic macros raw response (no macros parsed): {preview}");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: FetchMosaicMacros — parse error: {ex.Message}");
            _mosaicMacrosFetched = false; // Allow retry
        }
    }

    private static void ParseMacroArray(JsonElement array, Dictionary<string, string> macros)
    {
        foreach (var item in array.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object) continue;

            // RadPair schema: "phrase" = trigger name, "full_text" = expansion
            // Also try common alternatives in case schema changes
            var name = GetStr(item, "phrase") ?? GetStr(item, "name") ?? GetStr(item, "title")
                ?? GetStr(item, "trigger") ?? GetStr(item, "keyword");
            var text = GetStr(item, "full_text") ?? GetStr(item, "content") ?? GetStr(item, "text")
                ?? GetStr(item, "expansion") ?? GetStr(item, "body");

            if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(text))
            {
                macros.TryAdd(name.Trim(), text.Trim());
            }
        }
    }

    // ═══════ PUBLIC API: TEMPLATE CORRECTION ═══════

    /// <summary>
    /// Set the study type by searching the combobox. Replaces the UIA study-selector dance.
    /// </summary>
    public bool SetStudyType(string searchText)
    {
        return SetStudyType(searchText, null);
    }

    /// <summary>
    /// Set the study type by searching the combobox and clicking the best-matching dropdown option.
    /// When studyDescription is provided, options are scored by body part + modality matching.
    /// </summary>
    public bool SetStudyType(string searchText, string? studyDescription)
    {
        if (!IsIframeConnected) return false;
        var escaped = JsonSerializer.Serialize(searchText);

        // Step 1: Focus combobox, clear it, type search text using React-compatible setter
        // React overrides HTMLInputElement.prototype.value setter — we must use the native one
        // to bypass React's tracking, then dispatch an 'input' event so React picks up the change.
        var jsType = $@"(() => {{
            const combos = document.querySelectorAll('[role=""combobox""]');
            if (combos.length === 0) return 'no_combo';
            const combo = combos[0];
            combo.focus();
            const input = combo.querySelector('input') || combo;
            if (input.tagName !== 'INPUT') return 'no_input';

            input.focus();
            const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
            nativeSetter.call(input, '');
            input.dispatchEvent(new Event('input', {{ bubbles: true }}));
            nativeSetter.call(input, {escaped});
            input.dispatchEvent(new Event('input', {{ bubbles: true }}));
            return 'ok';
        }})()";
        var typeResult = ExtractResultValue(SendToIframe(jsType));
        if (typeResult != "ok")
        {
            Logger.Trace($"CDP: SetStudyType step 1 failed: {typeResult}");
            return false;
        }

        // Step 2: Wait for Mosaic to search/filter, then read dropdown options
        Thread.Sleep(1500);

        // Read all option texts from dropdown
        const string jsReadOptions = @"(() => {
            const options = document.querySelectorAll('[role=""option""]');
            if (options.length === 0) return 'waiting';
            const texts = [];
            options.forEach(o => texts.push(o.textContent?.trim() || ''));
            return JSON.stringify(texts);
        })()";

        string? optionsJson = null;
        for (int i = 0; i < 6; i++) // up to ~6 more seconds
        {
            var result = ExtractResultValue(SendToIframe(jsReadOptions));
            if (result != null && result != "waiting")
            {
                optionsJson = result;
                break;
            }
            Thread.Sleep(1000);
        }

        if (optionsJson == null)
        {
            Logger.Trace("CDP: SetStudyType timed out waiting for dropdown options");
            return false;
        }

        // Step 3: Find best matching option
        int bestIndex = 0; // Default to first option
        try
        {
            var options = JsonSerializer.Deserialize<string[]>(optionsJson);
            if (options == null || options.Length == 0)
            {
                Logger.Trace("CDP: SetStudyType no options parsed");
                return false;
            }

            Logger.Trace($"CDP: SetStudyType found {options.Length} options");

            if (!string.IsNullOrEmpty(studyDescription))
            {
                // Match using body parts + modality (same logic as UIA AttemptCorrectTemplate)
                var descParts = AutomationService.ExtractBodyParts(studyDescription);
                descParts.ExceptWith(AutomationService.OrganKeywords);
                descParts.Remove("CTA");
                descParts.Remove("MRA");
                var descModality = AutomationService.ExtractModality(studyDescription);

                for (int i = 0; i < options.Length; i++)
                {
                    Logger.Trace($"  [{i}] '{options[i]}'");
                    var itemParts = AutomationService.ExtractBodyParts(options[i]);
                    itemParts.ExceptWith(AutomationService.OrganKeywords);
                    itemParts.Remove("CTA");
                    itemParts.Remove("MRA");
                    var itemModality = AutomationService.ExtractModality(options[i]);
                    bool modalityMatch = descModality == null || itemModality == null
                        || string.Equals(descModality, itemModality, StringComparison.OrdinalIgnoreCase);
                    if (modalityMatch && itemParts.Count > 0 && descParts.SetEquals(itemParts))
                    {
                        bestIndex = i;
                        Logger.Trace($"  → Best match at [{i}]: modality+body parts match");
                        break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP: SetStudyType option matching error: {ex.Message}");
        }

        // Step 4: Click the matched option, then blur to prevent Mosaic re-searching
        var jsClickOption = $@"(() => {{
            const options = document.querySelectorAll('[role=""option""]');
            if (options.length <= {bestIndex}) return 'no_option';
            options[{bestIndex}].click();
            return 'ok';
        }})()";
        var clickResult = ExtractResultValue(SendToIframe(jsClickOption));
        Logger.Trace($"CDP: SetStudyType clicked option [{bestIndex}]: {clickResult}");

        if (clickResult == "ok")
        {
            // Brief pause for MUI to process the selection, then blur the input
            // to prevent the leftover search text from triggering a re-search
            Thread.Sleep(500);
            SendToIframe(@"(() => {
                const combos = document.querySelectorAll('[role=""combobox""]');
                if (combos.length > 0) {
                    const input = combos[0].querySelector('input') || combos[0];
                    if (input.blur) input.blur();
                }
                // Also click somewhere neutral to fully deselect
                const editors = document.querySelectorAll('.ProseMirror');
                if (editors.length > 0) editors[editors.length - 1].focus();
            })()");
        }

        return clickResult == "ok";
    }

    // ═══════ PUBLIC API: AUTO-SCROLL TO CURSOR ═══════

    /// <summary>
    /// Scroll the cursor into view in the given editor's scrollable area.
    /// Called after InsertContent (STT paste) to keep cursor visible during dictation.
    /// </summary>
    public void ScrollCursorIntoView(int editorIndex)
    {
        if (!IsIframeConnected || !AutoScrollEnabled) return;
        var js = $@"(() => {{
            const editors = document.querySelectorAll('.ProseMirror');
            const ed = editors[{editorIndex}];
            if (!ed) return 'no_editor';
            const area = ed.closest('[data-mt-editor-area]');
            if (!area) return 'no_area';
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0) return 'no_sel';
            const range = sel.getRangeAt(0);
            const rect = range.getBoundingClientRect();
            const areaRect = area.getBoundingClientRect();
            if (rect.bottom > areaRect.bottom - 20) {{
                area.scrollTop += (rect.bottom - areaRect.bottom + 60);
                return 'scrolled';
            }}
            if (rect.top < areaRect.top + 20) {{
                area.scrollTop -= (areaRect.top - rect.top + 60);
                return 'scrolled_up';
            }}
            return 'visible';
        }})()";
        try { SendToIframe(js); } catch { }
    }

    /// <summary>
    /// Inject a lightweight auto-scroll watcher that runs on a 300ms interval.
    /// Used for Dragon/built-in dictation where we can't hook into insert events.
    /// Checks if cursor is below viewport and scrolls it into view.
    /// </summary>
    public void InjectAutoScrollWatcher()
    {
        if (!IsIframeConnected || !AutoScrollEnabled || _autoScrollWatcherActive) return;
        var js = @"(() => {
            if (window.__mtAutoScrollInterval) return 'already_active';
            window.__mtAutoScrollInterval = setInterval(() => {
                const sel = window.getSelection();
                if (!sel || sel.rangeCount === 0) return;
                const range = sel.getRangeAt(0);
                const rect = range.getBoundingClientRect();
                if (rect.width === 0 && rect.height === 0) return; // collapsed/invisible
                // Find the editor area containing the selection
                let node = sel.anchorNode;
                while (node && !node.classList?.contains('ProseMirror')) node = node.parentElement;
                if (!node) return;
                const area = node.closest('[data-mt-editor-area]');
                if (!area) return;
                const areaRect = area.getBoundingClientRect();
                if (rect.bottom > areaRect.bottom - 20) {
                    area.scrollTop += (rect.bottom - areaRect.bottom + 60);
                } else if (rect.top < areaRect.top + 20) {
                    area.scrollTop -= (areaRect.top - rect.top + 60);
                }
            }, 300);
            return 'injected';
        })()";
        try
        {
            var result = ExtractResultValue(SendToIframe(js));
            _autoScrollWatcherActive = true;
            Logger.Trace($"CDP: Auto-scroll watcher: {result}");
        }
        catch { }
    }

    /// <summary>Remove the auto-scroll watcher interval.</summary>
    public void RemoveAutoScrollWatcher()
    {
        if (!_autoScrollWatcherActive) return;
        _autoScrollWatcherActive = false;
        if (!IsIframeConnected) return;
        var js = @"(() => {
            if (window.__mtAutoScrollInterval) {
                clearInterval(window.__mtAutoScrollInterval);
                window.__mtAutoScrollInterval = null;
                return 'removed';
            }
            return 'not_active';
        })()";
        try { SendToIframe(js); } catch { }
    }

    // ═══════ HELPERS ═══════

    private static string? GetStr(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var v) && v.ValueKind == JsonValueKind.String)
            return v.GetString();
        return null;
    }

    private static bool GetBool(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var v))
        {
            if (v.ValueKind == JsonValueKind.True) return true;
            if (v.ValueKind == JsonValueKind.False) return false;
        }
        return false;
    }

    // ═══════ DISPOSE ═══════

    // ═══════ PUBLIC API: CLARIO CDP ═══════

    /// <summary>
    /// Extract Priority, Class, and Accession from Clario via CDP JS eval.
    /// Returns null if not connected or fields not found.
    /// </summary>
    public ClarioPriorityResult? ClarioExtractPriorityAndClass(string? targetAccession = null)
    {
        if (!IsClarioConnected) return null;

        var js = @"(() => {
            const results = {};
            const labels = document.querySelectorAll('label, .x-form-item-label');
            for (const label of labels) {
                const text = label.textContent.trim().replace(/:$/, '');
                if (['Priority', 'Class', 'Accession'].includes(text)) {
                    const parent = label.closest('.x-form-item, .x-field, tr, div');
                    if (parent) {
                        const input = parent.querySelector('input, .x-form-display-field, .x-form-text, td:last-child');
                        if (input) {
                            results[text] = (input.value || input.textContent || '').trim();
                        }
                    }
                }
            }
            return JSON.stringify(results);
        })()";

        var raw = ExtractResultValue(SendToClario(js));
        if (raw == null) return null;

        try
        {
            var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;

            var priority = root.TryGetProperty("Priority", out var p) ? p.GetString() ?? "" : "";
            var cls = root.TryGetProperty("Class", out var c) ? c.GetString() ?? "" : "";
            var accession = root.TryGetProperty("Accession", out var a) ? a.GetString() ?? "" : "";

            // Verify accession if provided
            if (targetAccession != null && !string.IsNullOrEmpty(accession) &&
                !accession.Equals(targetAccession, StringComparison.OrdinalIgnoreCase))
            {
                Logger.Trace($"CDP Clario: Accession mismatch - expected '{targetAccession}', got '{accession}'");
                return null;
            }

            if (string.IsNullOrEmpty(priority) && string.IsNullOrEmpty(cls))
            {
                Logger.Trace("CDP Clario: No Priority or Class found");
                return null;
            }

            Logger.Trace($"CDP Clario: Priority='{priority}', Class='{cls}', Accession='{accession}'");
            return new ClarioPriorityResult(priority, cls, accession);
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP Clario: Parse error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Scrape exam note text from Clario via CDP JS eval.
    /// Returns the best (newest/longest) exam note text, or null if none found.
    /// </summary>
    public string? ClarioScrapeExamNote()
    {
        if (!IsClarioConnected) return null;

        var js = @"(() => {
            const notes = [];

            // Check textareas (most common container for exam notes)
            document.querySelectorAll('textarea').forEach(ta => {
                if (ta.value && ta.value.length > 10) {
                    notes.push(ta.value);
                }
            });

            // Check contenteditable elements
            document.querySelectorAll('[contenteditable=true]').forEach(ce => {
                const text = ce.textContent.trim();
                if (text.length > 10) {
                    notes.push(text);
                }
            });

            // Check for EXAM NOTE DataItem-like elements (ExtJS grid rows)
            document.querySelectorAll('.x-grid-cell-inner, .x-grid-row').forEach(el => {
                const text = el.textContent.trim();
                if (text.includes('EXAM NOTE') && text.length > 30) {
                    notes.push(text);
                }
            });

            // Check for note dialog if open
            const noteDialog = document.getElementById('content_patient_note_dialog_Main');
            if (noteDialog) {
                const noteField = noteDialog.querySelector('[id*=""noteFieldMessage""]');
                if (noteField) {
                    const text = noteField.value || noteField.textContent || '';
                    if (text.trim().length > 10) {
                        notes.unshift(text.trim()); // dialog note takes priority
                    }
                }
            }

            return JSON.stringify(notes);
        })()";

        var raw = ExtractResultValue(SendToClario(js));
        if (raw == null) return null;

        try
        {
            var notes = JsonSerializer.Deserialize<List<string>>(raw);
            if (notes == null || notes.Count == 0)
            {
                Logger.Trace("CDP Clario: No exam notes found");
                return null;
            }

            // Filter to plausible exam notes: must contain a date/time pattern or "EXAM NOTE"
            // This excludes status messages like "Opening exam(s) in VR..."
            var plausible = notes.Where(n =>
                n.Contains("EXAM NOTE", StringComparison.OrdinalIgnoreCase) ||
                System.Text.RegularExpressions.Regex.IsMatch(n, @"\d{1,2}[:/]\d{2}")).ToList();

            if (plausible.Count == 0)
            {
                Logger.Trace($"CDP Clario: Found {notes.Count} text(s) but none look like exam notes");
                return null;
            }

            // Return best match (dialog note if open, otherwise EXAM NOTE first, then longest)
            var best = plausible
                .OrderByDescending(n => n.Contains("EXAM NOTE", StringComparison.OrdinalIgnoreCase) ? 1 : 0)
                .ThenByDescending(n => n.Length)
                .First();

            Logger.Trace($"CDP Clario: Found exam note (len={best.Length})");
            return best;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CDP Clario: ExamNote parse error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Create a Critical Results Communication Note in Clario via CDP.
    /// Clicks: Create → Communication Note → Submit via ExtJS/DOM automation.
    /// </summary>
    public bool ClarioCreateCriticalNote()
    {
        if (!IsClarioConnected) return false;

        // Step 1: Find and click the "Create" button
        var jsClickCreate = @"(() => {
            // Try ExtJS button query first
            const buttons = document.querySelectorAll('.x-btn, a.x-btn, button, [role=""button""]');
            for (const btn of buttons) {
                const text = (btn.textContent || '').trim();
                if (text === 'Create' || text.startsWith('Create')) {
                    btn.click();
                    return 'clicked';
                }
            }
            // Try matching by inner span (ExtJS renders text in spans)
            const spans = document.querySelectorAll('.x-btn-inner, .x-btn-text');
            for (const span of spans) {
                if (span.textContent.trim() === 'Create') {
                    const btn = span.closest('.x-btn, a, button, [role=""button""]');
                    if (btn) { btn.click(); return 'clicked'; }
                    span.click();
                    return 'clicked_span';
                }
            }
            return 'not_found';
        })()";

        var createResult = ExtractResultValue(SendToClario(jsClickCreate));
        if (createResult == null || createResult == "not_found")
        {
            Logger.Trace($"CDP Clario: Create button not found (result={createResult})");
            return false;
        }

        Thread.Sleep(300); // Wait for menu to appear

        // Step 2: Find and click "Communication Note" menu item
        var jsClickCommNote = @"(() => {
            // Search visible menu items
            const items = document.querySelectorAll('.x-menu-item, .x-menu-item-text, [role=""menuitem""], .x-menu a');
            for (const item of items) {
                const text = (item.textContent || '').trim();
                if (text === 'Communication Note' || text.includes('Communication Note')) {
                    item.click();
                    return 'clicked';
                }
            }
            // Broader search for any clickable element with that text
            const all = document.querySelectorAll('a, span, div, li');
            for (const el of all) {
                if (el.children.length === 0 && el.textContent.trim() === 'Communication Note') {
                    el.click();
                    return 'clicked_text';
                }
            }
            return 'not_found';
        })()";

        var commResult = ExtractResultValue(SendToClario(jsClickCommNote));
        if (commResult == null || commResult == "not_found")
        {
            Logger.Trace($"CDP Clario: Communication Note menu item not found (result={commResult})");
            return false;
        }

        Thread.Sleep(400); // Wait for dialog to appear

        // Step 3: Find and click "Submit" button (retry up to 3 times)
        for (int retry = 0; retry < 3; retry++)
        {
            if (retry > 0) Thread.Sleep(200);

            var jsClickSubmit = @"(() => {
                const buttons = document.querySelectorAll('.x-btn, a.x-btn, button, [role=""button""]');
                for (const btn of buttons) {
                    const text = (btn.textContent || '').trim();
                    if (text === 'Submit') {
                        btn.click();
                        return 'clicked';
                    }
                }
                const spans = document.querySelectorAll('.x-btn-inner, .x-btn-text');
                for (const span of spans) {
                    if (span.textContent.trim() === 'Submit') {
                        const btn = span.closest('.x-btn, a, button, [role=""button""]');
                        if (btn) { btn.click(); return 'clicked'; }
                        span.click();
                        return 'clicked_span';
                    }
                }
                return 'not_found';
            })()";

            var submitResult = ExtractResultValue(SendToClario(jsClickSubmit));
            if (submitResult != null && submitResult != "not_found")
            {
                Logger.Trace("CDP Clario: CreateCriticalNote SUCCESS");
                return true;
            }
        }

        Logger.Trace("CDP Clario: Submit button not found after retries");
        return false;
    }

    /// <summary>Result of Clario Priority/Class extraction via CDP.</summary>
    public record ClarioPriorityResult(string Priority, string Class, string Accession);

    public void Dispose()
    {
        try
        {
            _slimHubCts?.Cancel();
            _iframeCts?.Cancel();
            _clarioCts?.Cancel();

            if (_slimHubWs?.State == WebSocketState.Open)
                try { _slimHubWs.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None).Wait(1000); } catch { }
            if (_iframeWs?.State == WebSocketState.Open)
                try { _iframeWs.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None).Wait(1000); } catch { }
            if (_clarioWs?.State == WebSocketState.Open)
                try { _clarioWs.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None).Wait(1000); } catch { }

            _slimHubWs?.Dispose();
            _iframeWs?.Dispose();
            _clarioWs?.Dispose();
            _slimHubCts?.Dispose();
            _iframeCts?.Dispose();
            _clarioCts?.Dispose();
        }
        catch { }

        // Complete any pending requests
        foreach (var kvp in _pending)
            kvp.Value.TrySetCanceled();
        _pending.Clear();

        Logger.Trace("CDP: Disposed");
    }
}

/// <summary>
/// Data class holding all scraped metadata from a CDP scrape operation.
/// Fields mirror AutomationService.Last* properties.
/// </summary>
public class CdpScrapeResult
{
    public string? PatientName { get; set; }
    public string? PatientGender { get; set; }
    public string? Mrn { get; set; }
    public string? SiteCode { get; set; }
    public string? Accession { get; set; }
    public string? Description { get; set; }
    public string? TemplateName { get; set; }
    public string? ReportText { get; set; }
    public int? PatientAge { get; set; }
    public bool IsDrafted { get; set; }
    public bool IsAddendum { get; set; }
}

/// <summary>
/// Structured report parsed from ProseMirror JSON tree.
/// Sections split at heading nodes, with dictated vs template text separation
/// based on Mosaic's highlight marks (highlighted = user-dictated, non-highlighted = template boilerplate).
/// IMPRESSION items are always treated as dictated.
/// </summary>
public class StructuredReport
{
    public List<ReportSection> Sections { get; set; } = new();

    /// <summary>Get a section by name (case-insensitive, strips trailing colon).</summary>
    public ReportSection? GetSection(string name)
    {
        var normalized = name.TrimEnd(':').Trim();
        return Sections.FirstOrDefault(s =>
            string.Equals(s.Name, normalized, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Get full report text (all sections concatenated).</summary>
    public string GetFullText() => string.Join("\n", Sections.Select(s => s.GetFullTextWithHeader()));

    /// <summary>Get only user-dictated text (non-highlighted, non-template).</summary>
    public string GetDictatedText() => string.Join("\n", Sections
        .Where(s => !string.IsNullOrWhiteSpace(s.DictatedText))
        .Select(s => s.Name + ":\n" + s.DictatedText));
}

public class ReportSection
{
    /// <summary>Section name without colon, e.g. "EXAM", "FINDINGS", "IMPRESSION".</summary>
    public string Name { get; set; } = "";

    /// <summary>All text in this section (template + dictated).</summary>
    public string FullText { get; set; } = "";

    /// <summary>Only highlighted (user-dictated) text. IMPRESSION items always included.</summary>
    public string DictatedText { get; set; } = "";

    /// <summary>Only non-highlighted (template boilerplate) text.</summary>
    public string TemplateText { get; set; } = "";

    /// <summary>For IMPRESSION: individual numbered items.</summary>
    public List<string> Items { get; set; } = new();

    /// <summary>For FINDINGS: subsections like LOWER CHEST, HEPATOBILIARY, etc.</summary>
    public List<ReportSection> Subsections { get; set; } = new();

    public string GetFullTextWithHeader() =>
        string.IsNullOrWhiteSpace(FullText) ? Name + ":" : Name + ":\n" + FullText;
}

# FlaUI/UIA Performance Fixes — Reference for RvuCounter

This document describes every performance bug discovered and fixed in MosaicTools' FlaUI-based UI Automation scraping of Chrome/Electron apps (Mosaic, Clario). RvuCounter uses many of the same patterns and likely has the same issues. Apply these fixes wherever the equivalent patterns exist.

---

## Critical Context: How UIA + Chrome Causes Progressive Degradation

Chrome-based apps (Chrome, Electron) expose their accessibility tree via UIA. Key facts:

1. **Chrome's accessibility tree NEVER shrinks** — Every `FindAllDescendants()` / `FindAllChildren()` call expands Chrome's internal accessibility node tracking. Nodes created for UIA queries are cached by Chrome indefinitely per-client connection. Over time, the tree grows larger and every subsequent query takes longer.

2. **UIA COM objects (RCWs) are expensive** — Each `AutomationElement` returned by FlaUI wraps a native `IUIAutomationElement` COM object via a Runtime Callable Wrapper (RCW). These hold cross-process references. If not explicitly released, they accumulate in the RCW table, making every subsequent UIA call progressively slower (the .NET RCW table is a global hashtable).

3. **Property reads are cross-process COM calls** — Every `.Name`, `.ControlType`, `.ClassName` read on an `AutomationElement` is a synchronous cross-process COM call to the target process (Chrome, Clario, etc.). In tight loops over hundreds of elements, this adds up to hundreds of COM round-trips per tick.

4. **UIA on STA thread blocks the UI** — `UIA3Automation` created on an STA thread causes all UIA calls from other threads to be marshaled through that thread's message pump. This blocks the UI thread during scraping.

---

## Fix 1: Release ALL AutomationElements After Use (THE MOST IMPORTANT FIX)

**Problem:** Every `FindFirst*()`, `FindAll*()`, `.Parent`, `.FindAllChildren()`, `.FindAllDescendants()` returns AutomationElement(s) that wrap COM objects. If not explicitly released, the COM wrappers accumulate indefinitely. .NET's GC will eventually finalize them, but not fast enough — the RCW table grows, and every subsequent UIA call gets slower.

**Pattern to look for:**
```csharp
// BAD — elements never released
var children = element.FindAllChildren();
foreach (var child in children)
{
    var name = child.Name;
    // ... use name ...
}
// children array and all its elements leak!

// BAD — single element never released
var found = parent.FindFirstDescendant(cf => cf.ByName("Something"));
if (found != null)
{
    var text = found.Name;
    // ... use text ...
}
// found leaks!
```

**Fix — always release in finally blocks:**
```csharp
// GOOD — array of elements
AutomationElement[]? children = null;
try
{
    children = element.FindAllChildren();
    foreach (var child in children)
    {
        var name = child.Name;
        // ... use name ...
    }
}
finally
{
    ReleaseElements(children);
}

// GOOD — single element
AutomationElement? found = null;
try
{
    found = parent.FindFirstDescendant(cf => cf.ByName("Something"));
    if (found != null)
    {
        var text = found.Name;
    }
}
finally
{
    ReleaseElement(found);
}
```

**The release helpers:**
```csharp
internal static void ReleaseElement(AutomationElement? el)
{
    if (el == null) return;
    try
    {
        if (el.FrameworkAutomationElement is UIA3FrameworkAutomationElement uia3)
            Marshal.ReleaseComObject(uia3.NativeElement);
    }
    catch (InvalidComObjectException) { }
    catch (Exception) { }
}

internal static void ReleaseElements(AutomationElement[]? elements)
{
    if (elements == null) return;
    foreach (var el in elements)
        ReleaseElement(el);
}
```

**Common leak sites to audit:**
- `FindAllChildren()` in loops (recursive tree walks are the worst — leak at every depth level)
- `FindAllDescendants()` results iterated but never released
- `FindFirstDescendant()` / `FindFirstChild()` results used and abandoned
- `.Parent` — returns a new AutomationElement that needs releasing
- `desktop.FindAllChildren(cf => cf.ByControlType(Window))` — window enumeration
- Any element kept across ticks as a "cache" that gets replaced without releasing the old one

**Recursive traversal example (worst leak pattern):**
```csharp
// BAD — leaks at every depth level
void Traverse(AutomationElement element, int depth)
{
    var children = element.FindAllChildren(); // LEAKED!
    foreach (var child in children)
        Traverse(child, depth + 1);
}

// GOOD
void Traverse(AutomationElement element, int depth)
{
    AutomationElement[]? children = null;
    try
    {
        // ... read properties of element ...
        children = element.FindAllChildren();
        foreach (var child in children)
            Traverse(child, depth + 1);
    }
    catch { }
    finally
    {
        ReleaseElements(children);
    }
}
```

---

## Fix 2: Create UIA3Automation on MTA Thread, NOT STA

**Problem:** If `new UIA3Automation()` is called on the UI thread (STA), the COM object (`CUIAutomation`, ThreadingModel=Both) lives in the STA apartment. All subsequent UIA calls from background threads get marshaled through the STA message pump, blocking the UI thread. This causes mouse lag, keyboard hook death, and general unresponsiveness during scraping.

**Fix:**
```csharp
// BAD
_automation = new UIA3Automation();

// GOOD — force MTA creation
_automation = Task.Run(() => new UIA3Automation()).GetAwaiter().GetResult();
```

`Task.Run` executes on the ThreadPool, which is MTA. The COM object lives in the MTA apartment, so background thread calls go directly to UIA without marshaling through the UI thread.

---

## Fix 3: Cache the Desktop Element

**Problem:** `_automation.GetDesktop()` creates a new COM wrapper every call. If called every scrape tick, it accumulates RCW references to the same underlying desktop element.

**Fix:**
```csharp
private AutomationElement? _cachedDesktop;

public AutomationElement GetCachedDesktop()
{
    var desktop = _cachedDesktop;
    if (desktop != null)
    {
        try
        {
            _ = desktop.Properties.ProcessId.ValueOrDefault; // Validate still alive
            return desktop;
        }
        catch
        {
            _cachedDesktop = null;
        }
    }
    desktop = _automation.GetDesktop();
    _cachedDesktop = desktop;
    return desktop;
}
```

---

## Fix 4: Don't Let ProseMirror/Document Search Errors Invalidate Window Caches

**Problem (the single biggest degradation driver we found):** When scraping Chrome-based apps, `FindAllDescendants()` on a document element can throw `E_FAIL` (HRESULT) when the document has no accessible content (e.g., no study loaded, page loading, etc.). If this exception propagates to an outer catch that invalidates the window cache, the next tick must re-find the window from scratch. This means:

1. Re-enumerate all desktop windows (`FindAllChildren` on desktop)
2. Re-search for the target document (`FindAllDescendants` with condition)
3. Re-attempt the content search that will fail again

That's 3 full subtree traversals per tick. Chrome's accessibility tree grows with each traversal. Over hours, ticks go from 500ms to 7000ms+.

**Fix — isolate content search errors from window/document cache invalidation:**
```csharp
// Outer method — finds window, caches it
public string? ScrapeReport()
{
    try
    {
        if (_cachedWindow == null)
            _cachedWindow = FindTargetWindow(); // expensive
        if (_cachedWindow == null) return null;

        return ScrapeReportInner(); // may throw on content search
    }
    catch (Exception ex)
    {
        // This invalidates the window cache — should ONLY fire for true window errors
        Logger.Trace($"Scrape error: {ex.Message}");
        InvalidateWindowCache();
        return null;
    }
}

// Inner method — content search wrapped in its own try/catch
private string? ScrapeReportInner()
{
    // ... find document, accession, etc. ...

    // Content search that may E_FAIL — catch locally!
    try
    {
        var candidates = document.FindAllDescendants(cf => cf.ByClassName("ProseMirror"));
        // ... process candidates ...
    }
    catch (Exception ex)
    {
        // Log but DON'T invalidate window/document caches!
        Logger.Trace($"Content search error: {ex.Message}");
    }

    return null; // No content found, but window/doc caches preserved
}
```

**Key principle:** The window and document are valid — only the report content is missing. Don't nuke everything just because the content search failed. Preserving caches means the next tick only needs 1 cheap validation read + 1 content search attempt, NOT 3 full subtree traversals.

---

## Fix 5: Periodic UIA Connection Reset

**Problem:** Even with proper element release, Chrome's accessibility tree tracking grows over time per-client connection. UIAutomationCore.dll also accumulates internal state.

**Fix — periodically dispose and recreate UIA3Automation:**
```csharp
private long _lastUiaResetTick64;
private int _uiaCallCount;
private const long UiaResetIntervalMs = 300_000; // 5 minutes
private const int UiaResetCallThreshold = 100;

public void ResetUiaConnectionIfNeeded()
{
    long elapsed = Environment.TickCount64 - _lastUiaResetTick64;
    if (elapsed < UiaResetIntervalMs || _uiaCallCount < UiaResetCallThreshold)
        return;

    Logger.Trace("Periodic UIA reset");

    // Release ALL cached elements first
    InvalidateAllCaches();
    ReleaseElement(_cachedDesktop);
    _cachedDesktop = null;

    // Force GC to release any remaining COM wrappers before disposing
    GC.Collect(2, GCCollectionMode.Forced);
    GC.WaitForPendingFinalizers();

    var callsSinceReset = Interlocked.Exchange(ref _uiaCallCount, 0);

    // Dispose old automation
    try { _automation.Dispose(); } catch { }

    // Create fresh automation (on current thread — should already be MTA if called from scrape thread)
    _automation = new UIA3Automation();

    _lastUiaResetTick64 = Environment.TickCount64;
    Logger.Trace($"UIA reset complete ({callsSinceReset} calls since last)");
}
```

Call this at the top of your scrape timer callback or at study change boundaries.

---

## Fix 6: Wrap All Find Results in try/finally (Even in Metadata/Sweep Loops)

**Problem:** Code that does `FindAllDescendants()` followed by an iteration loop, where the release code is AFTER the loop (not in a finally block). If any exception occurs during iteration, the elements are never released.

```csharp
// BAD — if el.ControlType throws, elements are never released
var elements = window.FindAllDescendants(condition);
foreach (var el in elements)
{
    var type = el.ControlType;  // could throw!
    var name = el.Name;         // could throw!
    // ... process ...
}
foreach (var el in elements)
    ReleaseElement(el);  // never reached if exception above
```

**Fix:**
```csharp
// GOOD
AutomationElement[]? elements = null;
try
{
    elements = window.FindAllDescendants(condition);
    foreach (var el in elements)
    {
        var type = el.ControlType;
        var name = el.Name;
        // ... process ...
    }
}
catch (Exception ex)
{
    Logger.Trace($"Sweep error: {ex.Message}");
}
finally
{
    ReleaseElements(elements);
}
```

---

## Fix 7: Rate-Limit FindAllDescendants Calls

**Problem:** `FindAllDescendants()` is the most expensive UIA operation — it traverses the entire subtree of the target element, creating COM wrappers for every node. On Chrome, it also expands Chrome's internal accessibility tree. Calling it every tick (2-3 seconds) is too aggressive.

**Fix — throttle with tick timestamps:**
```csharp
private long _lastFullSearchTick64;
private const long FullSearchMinIntervalMs = 10_000; // 10 seconds

// In scrape method:
long now = Environment.TickCount64;
if (now - _lastFullSearchTick64 < FullSearchMinIntervalMs)
{
    return _lastResult; // Return stale data
}
_lastFullSearchTick64 = now;
// ... do FindAllDescendants() ...
```

Different operations should have different intervals:
- **Metadata sweep** (drafted status, description): every 5 seconds
- **Full content search** (ProseMirror, report text): every 10 seconds
- **Patient info extraction** (unfiltered FindAllDescendants): only on study change + max 3 retries
- **Window enumeration**: only when cached window is invalid

---

## Fix 8: Async Logger (Don't Block Scrape Thread)

**Problem:** Synchronous file logging (`File.AppendAllText` or `StreamWriter.Write`) blocks the calling thread. If the scrape thread logs heavily, disk I/O stalls add to scrape tick time. Worse, if a keyboard hook callback logs something, the synchronous write can block long enough to trigger Windows' hook timeout (which silently removes the hook).

**Fix — queue-based async logger:**
```csharp
public static class Logger
{
    private static readonly ConcurrentQueue<string> _queue = new();
    private static readonly ManualResetEventSlim _signal = new(false);
    private static readonly Thread _writerThread;
    private static volatile bool _stopping;

    static Logger()
    {
        _writerThread = new Thread(WriterLoop)
        {
            IsBackground = true,
            Name = "LogWriter",
            Priority = ThreadPriority.BelowNormal
        };
        _writerThread.Start();
    }

    public static void Trace(string message)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        _queue.Enqueue($"{timestamp}: {message}\n");
        _signal.Set();
    }

    public static void Shutdown()
    {
        _stopping = true;
        _signal.Set();
        _writerThread.Join(3000);
    }

    private static void WriterLoop()
    {
        var sb = new StringBuilder();
        while (!_stopping)
        {
            _signal.Wait(1000);
            _signal.Reset();
            sb.Clear();
            while (_queue.TryDequeue(out var line))
                sb.Append(line);
            if (sb.Length > 0)
            {
                try { File.AppendAllText(_logPath, sb.ToString()); }
                catch { }
            }
        }
        // Flush remaining on shutdown
        var final = new StringBuilder();
        while (_queue.TryDequeue(out var line))
            final.Append(line);
        if (final.Length > 0)
            try { File.AppendAllText(_logPath, final.ToString()); } catch { }
    }
}
```

---

## Fix 9: Cache Elements Across Ticks (with Proper Invalidation)

**Problem:** Finding the target window, document, and editor every tick is expensive (multiple FindFirst/FindAll calls). Most ticks, these elements haven't changed.

**Fix — cache with validation:**
```csharp
private AutomationElement? _cachedWindow;
private AutomationElement? _cachedDocument;

// Validate cached window is still alive and correct
if (_cachedWindow != null)
{
    try
    {
        var title = _cachedWindow.Name?.ToLowerInvariant() ?? "";
        if (!title.Contains("expected-title"))
        {
            InvalidateWindowCache(); // releases element
        }
    }
    catch
    {
        InvalidateWindowCache(); // element died
    }
}

// Invalidation chain — release in order (deepest first)
private void InvalidateWindowCache()
{
    InvalidateDocumentCache();
    var win = _cachedWindow;
    _cachedWindow = null;
    ReleaseElement(win);
}

private void InvalidateDocumentCache()
{
    var doc = _cachedDocument;
    _cachedDocument = null;
    ReleaseElement(doc);
}

// When caching a new element, release the old one
private void CacheDocument(AutomationElement? newDoc)
{
    if (newDoc == null) return;
    var old = _cachedDocument;
    if (ReferenceEquals(old, newDoc)) return; // Same wrapper
    _cachedDocument = newDoc;
    ReleaseElement(old);
}
```

---

## Fix 10: Cap Retry Loops (Especially Recursive Traversals)

**Problem:** Retry loops that run `FindAllDescendants()` every tick until a condition is met can run indefinitely. A recursive depth-25 traversal running every 3 seconds leaks hundreds of COM wrappers per attempt.

**Fix:**
```csharp
private int _retryCount;
private const int MaxRetries = 5;

if (needsRetry && _retryCount < MaxRetries)
{
    _retryCount++;
    // ... attempt the expensive operation ...
}
else if (_retryCount >= MaxRetries)
{
    Logger.Trace("Gave up after max retries");
    _retryCount = 0; // Reset for next study
}
```

---

## Fix 11: Don't Create UIA3Automation Per Call

**Problem:** Creating a new `UIA3Automation()` for a one-off check is extremely expensive — it establishes a new COM connection to UIAutomationCore.dll, a new client ID, etc.

```csharp
// BAD — creates fresh UIA automation, uses it once, leaks it
public bool IsSomethingActive()
{
    using var automation = new UIA3Automation();
    var desktop = automation.GetDesktop();
    var window = desktop.FindFirstChild(cf => cf.ByName("Target"));
    return window != null;
}
```

**Fix:** Always reuse a shared `UIA3Automation` instance. If you need to call from a different context, pass the automation service in or use a singleton.

---

## Fix 12: FlaUI CacheRequest for Batch Property Reads (Optional Performance Boost)

**Problem:** Within loops over elements from `FindAllDescendants()`, reading `.Name`, `.ControlType`, `.ClassName` on each element is a separate cross-process COM call. For 200 elements × 2 properties = 400 COM calls.

**Fix — FlaUI CacheRequest pre-fetches properties with the Find call:**
```csharp
using FlaUI.Core;

private CacheRequest CreatePropertyCacheRequest()
{
    var cr = new CacheRequest();
    cr.TreeScope = TreeScope.Element | TreeScope.Children | TreeScope.Descendants;
    cr.Add(_automation.PropertyLibrary.Element.Name);
    cr.Add(_automation.PropertyLibrary.Element.ControlType);
    cr.Add(_automation.PropertyLibrary.Element.ClassName);
    return cr;
}

// Usage — wrap the scrape call:
using (CreatePropertyCacheRequest().Activate())
{
    return ScrapeInner(); // All Find* calls auto-route to BuildCache variants
}
// Within ScrapeInner(), .Name/.ControlType/.ClassName reads are local cache hits (0 COM calls)
```

**CRITICAL CAVEATS:**
- Elements cached from a PREVIOUS tick (outside the CacheRequest scope) will THROW `PropertyNotCachedException` when their properties are read within the scope. Use `CacheRequest.ForceNoCache()` to wrap reads on pre-cached elements:
  ```csharp
  using (CacheRequest.ForceNoCache())
  {
      var text = _cachedEditor.Name; // Forces live COM read, bypasses cache
  }
  ```
- Only include properties you actually read in loops. Reading a property NOT in the CacheRequest will throw.
- CacheRequest stacks are thread-local — safe for multi-threading.

---

## Fix 13: Heartbeat Logging for Diagnosing Leaks

Add periodic memory/state logging to detect degradation:

```csharp
// Every N scrape ticks (e.g., every 60 seconds):
Logger.Trace($"Scrape heartbeat: acc={accession}, idle={idleSeconds}");
Logger.Trace($"  Memory: managed={GC.GetTotalMemory(false)/1024/1024}MB, " +
    $"workingSet={Process.GetCurrentProcess().WorkingSet64/1024/1024}MB, " +
    $"private={Process.GetCurrentProcess().PrivateMemorySize64/1024/1024}MB, " +
    $"handles={Process.GetCurrentProcess().HandleCount}, " +
    $"threads={Process.GetCurrentProcess().Threads.Count}, " +
    $"GC={GC.CollectionCount(0)}/{GC.CollectionCount(1)}/{GC.CollectionCount(2)}, " +
    $"uiaCalls={_uiaCallCount}");
```

**What to watch for:**
- `managed` growing over time = managed object leak (probably AutomationElement arrays held in lists/fields)
- `private` growing = native COM leak (unreleased RCWs)
- `handles` growing = handle leak
- `uiaCalls` between resets — should be roughly constant per interval

---

## Audit Checklist

For every FlaUI usage site in RvuCounter, verify:

- [ ] Every `FindAllChildren()` result is released in a `finally` block
- [ ] Every `FindAllDescendants()` result is released in a `finally` block
- [ ] Every `FindFirstChild()` / `FindFirstDescendant()` result is released in a `finally` block
- [ ] Every `.Parent` access result is released in a `finally` block
- [ ] Recursive traversals release children at every depth level
- [ ] `UIA3Automation` is created on MTA thread (via `Task.Run`)
- [ ] Single shared `UIA3Automation` instance (not created per-call)
- [ ] Desktop element is cached, not re-fetched per tick
- [ ] Window element is cached with validation, not re-found per tick
- [ ] Content search errors don't invalidate window/document caches
- [ ] `FindAllDescendants()` calls are rate-limited (not every tick)
- [ ] Retry loops have max attempt caps
- [ ] Logger doesn't block calling thread (async/queue-based)
- [ ] Periodic UIA reset (dispose + recreate automation) every ~5 minutes
- [ ] Elements replaced in cache fields release the old element

---

## Impact Numbers (from MosaicTools)

Before fixes:
- Scrape ticks degraded from 500ms → 7300ms over 2 hours (14x slower)
- Memory grew from 4MB → 21MB managed, 34MB → 69MB private
- Keyboard hooks died every 30-40 seconds
- System became unusable after ~4 hours

After fixes:
- Scrape ticks flat at ~520ms over 75+ minutes (and counting)
- Memory stable at 4MB managed, 31MB private
- Keyboard hooks stable (reinstalls only during idle)
- No progressive degradation observed

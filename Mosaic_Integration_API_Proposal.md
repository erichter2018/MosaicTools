Mosaic Integration API Proposal
================================

From:    Erik Richter (MosaicTools - Radiology Workflow Automation)
To:      Mosaic Development Team
Date:    February 2026
Subject: Ideas for Structured Integration Points - Would Love Your Input


================================================================================
What This Is
================================================================================

I've built a radiology workflow tool called MosaicTools that extends Mosaic
with features like prior study insertion, critical findings documentation,
series capture, macro auto-insertion, AI-generated impressions, and
speech-to-text. It's used daily by radiologists at our site and has been
really well received.

The catch is that all communication with Mosaic is done by scraping the UI
tree via Microsoft UI Automation and injecting text via clipboard paste with
simulated keystrokes. It works, but it's fragile -- I've had to maintain three
separate code paths just for extracting the study description across Mosaic
2.0.2, 2.0.3, and 2.0.4.

I'd love to discuss whether there's appetite for some kind of supported
integration surface -- even something lightweight -- that would make this less
brittle and less hacky. I know this is asking for real work on your end, so I
want to lay out exactly what I need and why, and I'm completely open to
whatever approach makes sense for your architecture.


================================================================================
1. A Note on Security
================================================================================

I want to flag this not as a criticism but as context for why a supported API
would be a win for everyone:

UI Automation Is Wide Open
--------------------------
- Microsoft UI Automation gives full read access to every text element in
  Mosaic's process tree -- not just the fields I need, but everything: patient
  lists, worklist data, anything in the UI.
- Any third-party tool using UIA against Mosaic has this same unrestricted
  access. There's no way for Mosaic to scope or audit what's being read.
- A structured API would let you control exactly what data is exposed and
  to whom.

Keystroke Injection Is Invisible
--------------------------------
- I simulate Alt+R, Alt+P, Alt+F, Ctrl+V and other hotkeys to trigger Mosaic
  actions.
- This is indistinguishable from user input -- Mosaic can't tell whether a
  radiologist pressed Alt+P or my tool did.
- A command API would let you log integrator-initiated actions separately,
  which would be better for audit trails.

No Authentication
-----------------
- UI Automation requires zero authentication. Any process running as the same
  user can scrape Mosaic.
- A local API could use a shared secret, API key, or named-pipe ACL to
  restrict access to authorized integrators.

COM Wrapper Accumulation
------------------------
- FlaUI creates COM wrappers for every UI element traversed. These accumulate
  in my process's memory and I've already hit progressive system slowdown from
  the leaks. I've had to add forced GC cycles, retry caps, and idle backoff
  to manage it.

I'm not suggesting Mosaic has a security problem -- these are inherent
limitations of the scraping approach that a proper integration point would
naturally solve.


================================================================================
2. What I Currently Read from Mosaic
================================================================================

Here's everything I'm pulling out of the UI tree, along with how fragile each
extraction is. This should give you a clear picture of what a read API would
need to cover.

2.1 Report Content
------------------

Full report text
  How:       Scan ProseMirror elements in UIA tree, score by keyword density
             (TECHNIQUE, FINDINGS, IMPRESSION, etc.)
  Frequency: Every 2-30 seconds
  Notes:     Most expensive operation. Multiple ProseMirror elements exist
             (report vs. transcript); I distinguish them by U+FFFC object
             replacement character density.

Individual report sections
  How:       Regex parsing of full report text
  Frequency: On demand
  Notes:     I split on headers like IMPRESSION:, FINDINGS:, COMPARISON:, etc.
             This is error-prone.

Transcript text
  How:       Identify ProseMirror element WITHOUT U+FFFC characters
  Frequency: On demand
  Notes:     Used to detect transcript vs. final report content.

What would help -- something like:

    GET /report/current
    {
      "accession": "12345678",
      "sections": {
        "exam": "CT ABDOMEN PELVIS WITH IV CONTRAST",
        "clinical_history": "Abdominal pain, rule out appendicitis",
        "comparison": "CT abdomen pelvis 01/15/2026",
        "technique": "Axial images were obtained...",
        "findings": "The liver is normal in size...",
        "impression": "1. No acute abnormality.\n2. Stable appearance."
      },
      "full_text": "EXAM: CT ABDOMEN PELVIS...",
      "transcript_text": "The liver is normal in size..."
    }

Getting report content broken into named sections would be a huge win -- it'd
eliminate a whole class of regex fragility.


2.2 Study Metadata
------------------

Accession number
  How:       Find "Current Study" label, read next text element
  Fragility: Medium

Study description
  How:       Three version-specific code paths (inline text, separate
             label+button, inline button)
  Fragility: Very High -- this is the one that breaks every update

Patient name
  How:       Pattern match against exclusion list in UIA element text
  Fragility: High

Patient gender
  How:       Regex MALE|FEMALE in demographic text
  Fragility: Medium

Patient age
  How:       Regex AGE\s*(\d+) in same element
  Fragility: Medium

MRN
  How:       Version-dependent label detection
  Fragility: High

Site code
  How:       Positional extraction near accession
  Fragility: Medium

Template name
  How:       2nd line after EXAM: header in report
  Fragility: Low

What would help:

    GET /study/current
    {
      "accession": "12345678",
      "mrn": "MRN123456",
      "patient_name": "DOE, JOHN",
      "patient_age": 65,
      "patient_gender": "Male",
      "description": "CT ABDOMEN PELVIS WITH IV CONTRAST",
      "site_code": "MLC",
      "template_name": "CT ABD PELVIS",
      "priority": "STAT",
      "referring_physician": "Smith, Jane MD"
    }

Honestly, even just this one endpoint would eliminate most of my
version-specific scraping code.


2.3 Editor State
----------------

Drafted/signed status
  How:   Scan for "DRAFTED" label in UIA tree
  Notes: Triggers macro insertion and impression display

Dictation active
  How:   Windows registry mic consent store OR UIA tree scan
  Notes: Two fallback methods because neither is fully reliable on its own

Addendum mode
  How:   Check if report text starts with "Addendum"
  Notes: I need to block paste operations during addendums

Editor rebuild in progress
  How:   Compare pre/post Process Report text, wait for stabilization
  Notes: 15-25 second window after Process Report where the editor is
         rebuilding and scrapes return stale data

What would help:

    GET /editor/state
    {
      "status": "drafting",         // "empty"|"drafting"|"drafted"|"signed"|"addendum"
      "is_dictating": true,
      "is_processing": false,       // true during Process Report rebuild
      "active_editor": "transcript" // "transcript"|"final_report"|null
    }


2.4 Event Notifications -- The Big One
---------------------------------------

The single biggest improvement would be push-based notifications instead of
polling. Right now I'm scraping Mosaic's entire UI tree every 2-30 seconds
just to detect state changes. This is the primary cause of COM wrapper
accumulation and the progressive system slowdown I mentioned.

Events that would let me kill the polling loop entirely:

study.opened
  Payload:  accession, mrn, description, demographics
  Replaces: Accession change detection polling

study.closed
  Payload:  accession
  Replaces: Accession-empty debounce logic (I currently wait 3 poll ticks
            to confirm a study actually closed vs. a UI flap)

report.drafted
  Payload:  accession
  Replaces: Drafted state polling

report.signed
  Payload:  accession
  Replaces: Post-sign detection

report.changed
  Payload:  section, new_text
  Replaces: Full report re-scrape

dictation.started / dictation.stopped
  Payload:  --
  Replaces: Registry + UIA dictation state polling

editor.processing
  Payload:  started/completed
  Replaces: Post-Process-Report rebuild detection and timeout logic

If I could subscribe to even a few of these, the scrape timer could go away or
become much less aggressive, which would be a major improvement in system
resource usage.


================================================================================
3. What I Currently Write into Mosaic
================================================================================

3.1 How It Works Today (and Why It's Painful)
----------------------------------------------

Every text insertion follows this pattern:
  1. Save current window focus
  2. Activate Mosaic window (aggressive retry loop with Win32 thread attachment)
  3. Focus the correct editor (transcript vs. final report -- distinguished
     by U+FFFC character density, which is very hacky)
  4. Place text on system clipboard
  5. Send Ctrl+V to paste
  6. Wait for paste stabilization
  7. Restore previous focus

This is unreliable because:
- Focus management fails with multiple monitors, virtual desktops, or UAC
  dialogs
- Clipboard races -- if the user or another app touches the clipboard between
  my SetText and Ctrl+V, wrong content gets pasted
- Editor targeting is fragile -- the U+FFFC heuristic could break any time
- No feedback -- I have no way to confirm the paste actually succeeded
- Focus stealing -- the user sees Mosaic briefly flash to the foreground
  during every paste


3.2 Section-Level Insertion
---------------------------

What I'd really love is the ability to insert or replace content in a specific
report section without touching focus or clipboard:

    POST /report/section
    {
      "accession": "12345678",        // safety: ensure correct study
      "section": "impression",
      "action": "replace",            // "replace"|"append"|"prepend"
      "content": "1. No acute abnormality.\n2. Stable appearance."
    }

    POST /report/section
    {
      "accession": "12345678",
      "section": "clinical_history",
      "action": "append",
      "content": "\n\nCritical findings discussed with Dr. Smith at 3:45 PM on 02/18/2026."
    }

Sections I'd use:

impression
  Use case: AI-generated impression, RecoMD recommendations
  Current:  Click "IMPRESSION" header, Ctrl+Shift+End to select, Ctrl+V

comparison
  Use case: Prior study text
  Current:  Paste into transcript editor

transcript (full)
  Use case: Prior studies, macros, pick list items, STT text
  Current:  Focus transcript ProseMirror, Ctrl+V

final_report (full)
  Use case: Critical findings documentation
  Current:  Focus final report ProseMirror, Ctrl+V

clinical_history
  Use case: Not currently writing here, but would be useful

findings
  Use case: Not currently writing here, but would be useful


3.3 Editor Commands
-------------------

    POST /editor/command
    {
      "command": "process_report"   // "process_report"|"sign_report"|"toggle_dictation"
    }

This would replace my simulated Alt+P, Alt+F, Alt+R keystrokes and give you
the ability to:
- Log these as integration-initiated actions (better audit trail)
- Return success/failure status
- Reject commands when they don't make sense (e.g., signing during addendum)


3.4 Cursor-Position Insertion
-----------------------------

For speech-to-text, I need to insert text at the current cursor position
without stealing focus or using the clipboard:

    POST /editor/insert
    {
      "text": "The liver is normal in size and contour. ",
      "at": "cursor"
    }

This is my highest-frequency write -- real-time transcription inserts text
roughly every 200ms. It'd benefit the most from not having to round-trip
through the clipboard.


================================================================================
4. Transport -- I'm Flexible
================================================================================

I genuinely don't have a strong preference here. Whatever fits best with
Mosaic's architecture:

Local HTTP (localhost)
  Pros: Language-agnostic, easy to debug, WebSocket upgrade for events
  Cons: Needs port management

Named Pipes
  Pros: No network stack, Windows ACL for auth, very fast
  Cons: Windows-only (which is fine for us)

Window Messages (WM_COPYDATA)
  Pros: Zero dependencies, native Win32 IPC
  Cons: Size limits, no streaming

Shared Memory + Events
  Pros: Extremely fast, zero-copy reads
  Cons: Complex to synchronize and version

If you already have an IPC mechanism internally, piggybacking on that would
probably be the path of least resistance.


================================================================================
5. Incremental Path -- No Need to Boil the Ocean
================================================================================

I know building a full API is a lot of work, and I want to be respectful of
that. Here's a phased approach where each step is independently valuable:

Phase 1 -- Read-Only Study Metadata
  Expose accession, MRN, description, demographics, and report status.
  This alone would eliminate my most fragile code -- the three
  version-specific description extraction paths and the patient info retry
  loops.

Phase 2 -- Event Notifications
  Push notifications for study open/close and report state changes.
  This kills the polling loop, which is the biggest source of resource
  consumption.

Phase 3 -- Structured Report Sections
  Report content returned as named sections instead of raw text.
  Enables reliable section-targeted reads and writes.

Phase 4 -- Write API
  Programmatic text insertion into specific sections or at cursor.
  Eliminates clipboard usage and focus stealing.

Phase 5 -- Command API
  Trigger Process Report, Sign, Toggle Dictation programmatically.
  Better audit trail, proper error handling.

Even Phase 1 alone would save me significant maintenance headaches on every
Mosaic update.


================================================================================
6. What I'd Gain (and What You'd Gain)
================================================================================

Reliability
  Now:  Breaks with UI changes; 3 version-specific code paths today
  API:  Stable contract, versioned endpoints

Performance
  Now:  Polling every 2-30s, COM leaks, forced GC cycles
  API:  Event-driven, zero polling overhead

Security
  Now:  Unrestricted UIA access, unauditable keystrokes
  API:  Scoped access, auditable commands

Your maintenance burden
  Now:  I reverse-engineer new UI structure after every update to keep up
  API:  Clean integration contract that survives UI changes

User experience
  Now:  Focus stealing during paste, clipboard clobbering
  API:  Background operation, invisible to user


================================================================================
7. The Bigger Picture -- Enabling Innovation Without Risk
================================================================================

The reality of radiology workflows is that every site has its own quirks,
its own PACS, its own IT constraints, and its own ideas about what would make
their radiologists faster. A supported integration surface wouldn't just help
me -- it would give any motivated site the ability to build tooling that
addresses their unique needs without you having to anticipate every use case
or take on the risk of shipping niche features into the core product.

The features I've built exist because radiologists at my site had specific
pain points that were too site-specific to justify a product feature request.
An integration API would let that kind of innovation happen safely and
independently -- users can experiment and iterate on their own workflows while
Mosaic stays stable. And when something proves out broadly useful, it gives
you a clear signal about what to bring into the product itself.

In short: a query/command interface doesn't compete with Mosaic's roadmap --
it feeds it, while keeping the core application clean.


================================================================================
8. What I Can Offer
================================================================================

I'm happy to:
- Adopt anything incrementally -- even a single metadata endpoint would make
  a real difference
- Adapt to whatever transport you prefer -- HTTP, pipes, window messages,
  COM, whatever
- Test pre-release builds so we catch integration issues before they hit
  production
- Share my scraping code so you can see exactly which UI elements I depend on
  and why -- the tool is open source at github.com/erichter2018/MosaicTools
- Write the client library -- I don't need you to build a client SDK, just
  the endpoint

I really appreciate you taking the time to read this. Even if a full API isn't
feasible right now, I'd love to hear if there's a lighter-weight option I'm
not thinking of, or if there's anything on the Mosaic roadmap that would help
with third-party integration.

--------------------------------------------------------------------------------

Erik Richter -- erik.richter@radpartners.com
MosaicTools: github.com/erichter2018/MosaicTools

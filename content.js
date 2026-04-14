// MeetMark content script.
// Injected into a Microsoft Teams transcript page, or a SharePoint Stream
// recording page (stream.aspx) where Teams meeting recordings are served,
// on demand. Responsible for:
//   1) scrolling the page so all lazy-loaded transcript chunks render
//   2) scraping the DOM for speaker / timestamp / utterance triples
//   3) cleaning up known Teams text artifacts
//   4) formatting the result as Markdown
//   5) returning the markdown + filename to the popup
//
// Nothing here makes a network call. Everything is local to the tab.

(function () {
  // Guard against double injection. The popup may inject this script more than
  // once across multiple clicks; we want only one message listener active.
  if (window.__meetmarkLoaded) {
    return;
  }
  window.__meetmarkLoaded = true;

  // Maximum total time we are willing to spend scrolling for lazy content,
  // in milliseconds. Long meetings (2+ hours) can take a while to load.
  const MAX_SCROLL_MS = 60000;

  // Pause between incremental scroll steps. Teams renders virtualized
  // chunks; too fast and we skip rendering of newly-injected rows, causing
  // the lazy-load mechanism to reset. 500 ms gives each batch of DOM nodes
  // time to settle before the next scroll fires.
  const SCROLL_STEP_DELAY_MS = 500;

  // After we hit the bottom we wait this long and re-check whether the page
  // grew. If it didn't, we proceed with what we have.
  const SETTLE_DELAY_MS = 800;

  // Listen for port connections from the popup. The popup opens a long-lived
  // Port named "meetmark", sends { type: "start" } to begin, and listens for
  // { type: "progress" / "done" / "error" } messages back. It may also send
  // { type: "cancel" } at any time.
  chrome.runtime.onConnect.addListener((port) => {
    if (port.name !== "meetmark") return;

    let cancelled = false;

    port.onDisconnect.addListener(() => {
      cancelled = true;
    });

    port.onMessage.addListener((msg) => {
      if (!msg) return;
      if (msg.type === "cancel") {
        cancelled = true;
        return;
      }
      if (msg.type !== "start") return;

      function isCancelled() {
        return cancelled;
      }

      runExport(port, isCancelled)
        .then((result) => {
          if (cancelled) return;
          if (result.ok) {
            try {
              port.postMessage({
                type: "done",
                filename: result.filename,
                markdown: result.markdown,
                warning: result.warning || null,
                turnCount: result.turnCount || 0,
                speakerCount: result.speakerCount || 0,
              });
            } catch (_) {}
          } else {
            try {
              port.postMessage({
                type: "error",
                message: result.error || "Export failed.",
              });
            } catch (_) {}
          }
        })
        .catch((err) => {
          if (cancelled) return;
          try {
            port.postMessage({
              type: "error",
              message: (err && err.message) || String(err),
            });
          } catch (_) {}
        });
    });
  });

  // Top-level orchestration. Returns { ok, markdown, filename, warning,
  // turnCount, speakerCount } or { ok: false, error }.
  async function runExport(port, isCancelled) {
    function progress(message, detail) {
      if (!port) return;
      try {
        port.postMessage({ type: "progress", message, detail: detail || "" });
      } catch (_) {}
    }

    progress("Locating transcript panel...");

    // If the transcript panel isn't open yet, try to click the Transcript
    // button in the Stream player sidebar. On SharePoint Stream the URL does
    // not change when the panel opens, so this is the only hint we have.
    let container = findTranscriptContainer();
    if (!container) {
      const opened = await ensureTranscriptPanelOpen();
      if (opened) {
        container = findTranscriptContainer();
      }
    }
    if (!container) {
      return {
        ok: false,
        error:
          "Could not find transcript content. Open the Teams meeting recording in SharePoint Stream, click 'Transcript' to show the transcript panel, let it load, then try again.",
      };
    }

    if (isCancelled && isCancelled()) {
      return { ok: false, error: "Cancelled." };
    }

    progress("Reading transcript...");

    // Collect turns while scrolling. Stream uses a virtualized list, so rows
    // outside the viewport are removed from the DOM; we must scrape at every
    // scroll step and dedupe, not just scrape once at the end.
    let collectResult;
    try {
      collectResult = await scrollAndCollectTurns(container, port, isCancelled);
    } catch (err) {
      return {
        ok: false,
        error: "Scroll/scrape error: " + ((err && err.message) || String(err)),
      };
    }

    if (isCancelled && isCancelled()) {
      return { ok: false, error: "Cancelled." };
    }

    progress("Processing transcript...");

    const warning = collectResult.timedOut
      ? "Scroll timed out before all content loaded. Exporting what was collected."
      : null;

    const rawTurns = collectResult.turns || [];
    if (rawTurns.length === 0) {
      return {
        ok: false,
        error:
          "Could not find transcript content. Open the Teams meeting recording in SharePoint Stream, click 'Transcript' to show the transcript panel, let it load, then try again.",
      };
    }

    const cleanedTurns = rawTurns
      .map(cleanTurn)
      .filter((t) => t.text.length > 0 || t.speaker.length > 0);
    const mergedTurns = mergeAdjacentTurns(cleanedTurns);
    const participants = collectParticipants(mergedTurns);
    const metadata = extractMeetingMetadata();
    const markdown = formatMarkdown(mergedTurns, participants, metadata);
    const filename = buildFilename(metadata);

    return {
      ok: true,
      markdown,
      filename,
      warning,
      turnCount: mergedTurns.length,
      speakerCount: participants.length,
    };
  }

  // ---------------------------------------------------------------------------
  // Selector resolution
  // ---------------------------------------------------------------------------

  // Try a list of selectors in order and return the first matching element.
  // Logs which selector matched so it's easy to debug after a Teams update.
  function querySelectorWithFallbacks(root, selectors, label) {
    for (const sel of selectors) {
      const found = root.querySelector(sel);
      if (found) {
        console.log("[MeetMark] " + label + " matched selector: " + sel);
        return found;
      }
    }
    return null;
  }

  // Same as above but returns all matches under the first selector that hits.
  function querySelectorAllWithFallbacks(root, selectors, label) {
    for (const sel of selectors) {
      const found = root.querySelectorAll(sel);
      if (found && found.length > 0) {
        console.log(
          "[MeetMark] " +
            label +
            " matched selector: " +
            sel +
            " (" +
            found.length +
            " nodes)"
        );
        return found;
      }
    }
    return [];
  }

  // Locate the scroll-container that holds the transcript. Teams and
  // SharePoint Stream reuse a handful of class names depending on the
  // surface, so we try several. The Stream page (stream.aspx) is where most
  // tenants serve Teams meeting recordings and their transcripts.
  function findTranscriptContainer() {
    const containerSelectors = [
      // SharePoint Stream (stream.aspx) — known stable data-automationid hooks.
      '[data-automationid="transcript-virtualized-list"]',
      '[data-automationid="transcriptList"]',
      '[data-automationid="transcript-list"]',
      '[data-automationid="transcript-panel"] [role="list"]',
      '[data-automationid="transcript-panel"]',

      // Teams web recording-playback sidebar (teams.microsoft.com).
      // The sidebar panels use data-tid identifiers and Fluent UI roles.
      '[data-tid="calling-transcript-list"]',
      '[data-tid="calling-transcript"]',
      '[data-tid="recording-transcript-list"]',
      '[data-tid="recording-transcript-panel"]',
      '[data-tid="recording-transcript"]',
      '[data-tid="transcript-content"]',
      '[data-tid="transcript-list"]',
      '[data-tid="transcriptList"]',
      '[data-tid="closed-caption-renderer"]',

      // Role + aria-label combos (language-independent case-insensitive).
      '[role="list"][aria-label*="ranscript" i]',
      '[role="feed"][aria-label*="ranscript" i]',
      '[role="log"][aria-label*="ranscript" i]',
      '[role="region"][aria-label*="ranscript" i]',
      '[role="tabpanel"][aria-label*="ranscript" i]',
      'div[aria-label*="ranscript" i]',
      'section[aria-label*="ranscript" i]',
      'div[aria-label="Transcript"] [role="list"]',
      'div[aria-label="Transcript"]',

      // Broad class-name patterns Teams components often share.
      '[class*="calling-transcript"]',
      '[class*="meeting-transcript"]',
      '[class*="transcript-panel"]',
      '[class*="transcript-list"]',
      '[class*="ts-transcript"]',
      ".transcript-list",
      ".transcript",

      // Tab-panel containers — Teams recording sidebar wraps content in a
      // tabpanel element whose child list has the actual items.
      '[role="tabpanel"] [role="list"]',
      '[role="tabpanel"] [role="feed"]',
      '[role="tabpanel"]',
    ];

    const found = querySelectorWithFallbacks(
      document,
      containerSelectors,
      "transcript container"
    );
    if (found) return found;

    // Last-resort heuristic: the transcript container is the deepest DOM
    // element that contains at least three timestamp-like strings (e.g.
    // "0:03", "7:24"). Transcript panels are uniquely timestamp-dense;
    // no other sidebar panel looks like this.
    return findContainerByTimestampDensity();
  }

  // Walk visible block-level elements and return the deepest one that
  // contains at least MIN_STAMPS distinct timestamp strings. "Deepest" means
  // highest ancestor-count, which gives us the most-specific container.
  function findContainerByTimestampDensity() {
    const MIN_STAMPS = 3;
    const TIME_PAT = /\b\d{1,3}:\d{2}\b/g;
    const seen = new Set();
    const candidates = document.querySelectorAll(
      [
        "div", "section", "aside", "main", "ul", "ol",
        '[role="list"]', '[role="feed"]', '[role="log"]',
        '[role="region"]', '[role="tabpanel"]',
      ].join(",")
    );

    let best = null;
    let bestDepth = -1;

    for (const el of candidates) {
      if (el === document.body || el === document.documentElement) continue;
      if (seen.has(el)) continue;
      seen.add(el);

      // Skip elements that are hidden or zero-size.
      if (!el.offsetParent && el.tagName !== "BODY") {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 && rect.height === 0) continue;
      }

      const text = el.innerText || el.textContent || "";
      const stamps = text.match(TIME_PAT);
      if (!stamps || stamps.length < MIN_STAMPS) continue;

      // Count ancestor depth.
      let depth = 0;
      let node = el;
      while (node.parentElement) {
        depth++;
        node = node.parentElement;
      }

      if (depth > bestDepth) {
        best = el;
        bestDepth = depth;
      }
    }

    if (best) {
      console.log(
        "[MeetMark] transcript container found by timestamp density heuristic"
      );
    }
    return best;
  }

  // ---------------------------------------------------------------------------
  // Transcript panel auto-open
  // ---------------------------------------------------------------------------

  // If the user is on a SharePoint Stream page but has not clicked the
  // "Transcript" button yet, the transcript panel isn't in the DOM. Look for
  // the sidebar button by aria-label and click it. Returns true if the panel
  // appears (or was already open) within a few seconds.
  async function ensureTranscriptPanelOpen() {
    if (findTranscriptContainer()) return true;

    const buttonSelectors = [
      'button[aria-label="Transcript"][role="menuitem"]',
      'button[aria-label="Transcript"]',
      'button[aria-label*="ranscript"][role="menuitem"]',
      'button[aria-label*="ranscript"]',
      '[role="menuitem"][aria-label="Transcript"]',
    ];

    let clicked = false;
    for (const sel of buttonSelectors) {
      const btn = document.querySelector(sel);
      if (btn) {
        console.log("[MeetMark] clicking transcript button: " + sel);
        try {
          btn.click();
          clicked = true;
          break;
        } catch (err) {
          console.warn("[MeetMark] click failed: " + err);
        }
      }
    }

    if (!clicked) {
      console.log("[MeetMark] no transcript button found");
      return false;
    }

    // Poll for the panel to materialize.
    for (let i = 0; i < 24; i++) {
      await sleep(250);
      if (findTranscriptContainer()) return true;
    }
    return false;
  }

  // ---------------------------------------------------------------------------
  // Scrolling + incremental scraping
  // ---------------------------------------------------------------------------

  // Promise-based sleep helper.
  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  // Determine which element actually scrolls the transcript panel. Walk up
  // from the container looking first for an ancestor that declares an
  // explicit overflow-y style (auto/scroll/overlay) — these are almost
  // always the intended panel scroller. Fall back to the first generic
  // scrollable ancestor, and only as a last resort the page-level scroller.
  // We avoid defaulting to the page scroller because some surfaces reset it
  // whenever new virtualized content renders.
  function findScroller(start) {
    let node = start;
    let genericScrollable = null;
    while (node && node !== document.body) {
      if (node.scrollHeight > node.clientHeight + 4) {
        if (!genericScrollable) genericScrollable = node;
        const oy = window.getComputedStyle(node).overflowY;
        if (oy === "auto" || oy === "scroll" || oy === "overlay") {
          return node;
        }
      }
      node = node.parentElement;
    }
    // Some panels put the scroller INSIDE the list container rather than
    // around it. Check descendants for an explicit-overflow scroll host.
    const descendants = start.querySelectorAll("*");
    for (let i = 0; i < descendants.length && i < 500; i++) {
      const el = descendants[i];
      if (el.scrollHeight > el.clientHeight + 4) {
        const oy = window.getComputedStyle(el).overflowY;
        if (oy === "auto" || oy === "scroll" || oy === "overlay") {
          return el;
        }
      }
    }
    return (
      genericScrollable ||
      document.scrollingElement ||
      document.documentElement
    );
  }

  // Build a dedupe key for a turn. We key on speaker + timestamp + a prefix of
  // the text so that the same row captured by two different scroll positions
  // only ends up in the output once.
  function turnKey(turn) {
    const speaker = (turn.speaker || "").toLowerCase().trim();
    const ts = (turn.timestamp || "").trim();
    const textPrefix = (turn.text || "").slice(0, 80).toLowerCase().trim();
    return speaker + "|" + ts + "|" + textPrefix;
  }

  // Scroll incrementally from the top of the transcript to the bottom,
  // scraping visible rows at every step. Because the list is virtualized,
  // rows outside the viewport get removed from the DOM; scraping at every
  // step and deduping by { speaker, timestamp, text-prefix } is the only
  // way to collect the full transcript.
  //
  // Returns { turns, timedOut }.
  async function scrollAndCollectTurns(container, port, isCancelled) {
    let scroller = findScroller(container);
    const collected = new Map();
    const collectOrder = [];
    // Secondary dedup guard keyed on speaker+timestamp only. Catches cases
    // where the text differs slightly across scroll resets (e.g. partial
    // renders at the boundary) but the segment is the same logical utterance.
    const seenSegments = new Set();

    // Disable CSS scroll anchoring on the scroller (and, defensively, its
    // ancestor chain) while we scrape. Scroll anchoring is the browser-level
    // feature that repositions scrollTop to keep the currently-anchored
    // element visible when new content is inserted above it. In a virtualized
    // transcript list that injects rows above the viewport as you scroll,
    // this causes scrollTop to get snapped back toward 0 mid-export.
    const savedAnchors = [];
    function pinScrollAnchor(el) {
      if (!el || !el.style) return;
      savedAnchors.push({ el, prev: el.style.overflowAnchor });
      try {
        el.style.overflowAnchor = "none";
      } catch (_) {}
    }
    function restoreScrollAnchors() {
      for (const { el, prev } of savedAnchors) {
        try {
          el.style.overflowAnchor = prev || "";
        } catch (_) {}
      }
    }
    pinScrollAnchor(scroller);
    let anc = scroller && scroller.parentElement;
    let depth = 0;
    while (anc && depth < 6) {
      pinScrollAnchor(anc);
      anc = anc.parentElement;
      depth += 1;
    }

    // Push a turn into the collection, preserving first-seen order.
    function absorb(turn) {
      if (!turn) return;
      if (!turn.speaker && !turn.text) return;
      const segKey = (turn.speaker || "") + "|" + (turn.timestamp || "");
      if (segKey !== "|" && seenSegments.has(segKey)) return;
      const key = turnKey(turn);
      if (collected.has(key)) return;
      if (segKey !== "|") seenSegments.add(segKey);
      collected.set(key, turn);
      collectOrder.push(key);
    }

    // Scrape whatever is currently rendered in the panel.
    function scrapeOnce() {
      const turns = scrapeTurns(container);
      for (const t of turns) absorb(t);
    }

    // Re-find the scroller if the current one was detached from the DOM.
    // Teams/Stream occasionally swaps out the virtualized list's scroll host
    // when the user toggles the panel or the tab refocuses. Without this, the
    // stale reference silently no-ops on scrollTop writes and the scraper
    // thinks it's stuck.
    function refreshScrollerIfDetached() {
      if (!scroller || !scroller.isConnected) {
        const fresh = findScroller(container);
        if (fresh) {
          scroller = fresh;
          pinScrollAnchor(scroller);
        }
      }
    }

    const startTime = Date.now();

    // Start at the very top so rows render from the beginning of the meeting.
    scroller.scrollTop = 0;
    await sleep(SETTLE_DELAY_MS);
    scrapeOnce();

    let stableAtBottom = 0;
    let stableStuck = 0;
    let lastScrollTop = -1;
    let lastHeight = -1;
    let writeIgnoredCount = 0;

    function finish(timedOut) {
      try {
        scroller.scrollTop = 0;
      } catch (_) {}
      restoreScrollAnchors();
      return {
        turns: collectOrder.map((k) => collected.get(k)),
        timedOut,
      };
    }

    try {
      // The outer loop drives the scroll downward until we have been at the
      // very bottom, with stable scrollHeight, for several consecutive checks.
      // We deliberately require multiple stable confirmations because Stream's
      // virtualized list grows its scrollHeight as new rows are fetched.
      while (true) {
        if (isCancelled && isCancelled()) {
          return finish(false);
        }

        if (Date.now() - startTime > MAX_SCROLL_MS) {
          return finish(true);
        }

        refreshScrollerIfDetached();

        const beforeScroll = scroller.scrollTop;
        const beforeHeight = scroller.scrollHeight;
        const viewport = scroller.clientHeight || window.innerHeight || 600;
        const targetTop = beforeScroll + viewport * 0.7;

        // Advance by ~70% of viewport height so there is overlap between
        // scrape steps and no row gets skipped.
        scroller.scrollTop = targetTop;
        await sleep(SCROLL_STEP_DELAY_MS);
        scrapeOnce();

        const afterScroll = scroller.scrollTop;
        const afterHeight = scroller.scrollHeight;

        // Detect "write was ignored" — we tried to move forward but scrollTop
        // didn't actually advance AND the panel has more content below. This
        // usually means we're bound to the wrong element (e.g. the page
        // scroller). Re-resolve the scroller and try again before bailing.
        if (
          afterScroll <= beforeScroll + 1 &&
          afterHeight > (beforeScroll + viewport + 4)
        ) {
          writeIgnoredCount += 1;
          if (writeIgnoredCount === 2) {
            const fresh = findScroller(container);
            if (fresh && fresh !== scroller) {
              scroller = fresh;
              pinScrollAnchor(scroller);
            }
          }
        } else {
          writeIgnoredCount = 0;
        }

        // Report scroll progress so the popup can show something is happening.
        if (port && afterHeight > 0) {
          const pct = Math.min(100, Math.round((afterScroll / afterHeight) * 100));
          try {
            port.postMessage({
              type: "progress",
              message: "Reading transcript... " + pct + "%",
              detail: collectOrder.length + " turns collected",
            });
          } catch (_) {}
        }
        const atBottom =
          afterScroll + scroller.clientHeight >= afterHeight - 4;

        if (atBottom) {
          // Wait for any trailing lazy fetches to land, then re-scrape.
          await sleep(SETTLE_DELAY_MS);
          scrapeOnce();
          // Jump to the exact bottom to trigger any "almost there" fetches.
          scroller.scrollTop = scroller.scrollHeight;
          await sleep(SETTLE_DELAY_MS);
          scrapeOnce();

          if (scroller.scrollHeight === afterHeight) {
            stableAtBottom += 1;
            if (stableAtBottom >= 4) {
              // Truly done.
              return finish(false);
            }
          } else {
            stableAtBottom = 0;
          }
        } else {
          stableAtBottom = 0;
        }

        // Guard against a stuck scroller (e.g. transcript shorter than the
        // viewport). If neither the scrollTop nor the scrollHeight has changed
        // for several consecutive iterations, assume we're done.
        if (afterScroll === lastScrollTop && afterHeight === lastHeight) {
          stableStuck += 1;
          if (stableStuck >= 6) {
            return finish(false);
          }
        } else {
          stableStuck = 0;
          lastScrollTop = afterScroll;
          lastHeight = afterHeight;
        }
      }
    } finally {
      restoreScrollAnchors();
    }
  }

  // ---------------------------------------------------------------------------
  // Scraping
  // ---------------------------------------------------------------------------

  // Walk the container and pull out one record per transcript row.
  //
  // On SharePoint Stream the DOM classes vary between tenants and releases,
  // so we use a pattern-based scraper that walks visible text leaves and
  // groups them into turns based on the Name -> HH:MM -> Text pattern. If
  // that yields nothing we fall back to the class/attribute-driven scraper
  // which works for native Teams surfaces.
  function scrapeTurns(container) {
    // The pattern scraper walks text-leaf DOM elements in document order and
    // groups them by the Name → HH:MM → Text cadence that every Teams/Stream
    // transcript follows, regardless of class or attribute naming. Running it
    // first on every surface avoids the selector scraper's blind spot: it
    // matches continuation listitem rows (which carry text but no speaker
    // element) and emits them as speakerless "Unknown speaker" fragments.
    const patternTurns = scrapeTurnsByPattern(container);
    // Only trust the pattern scraper if it actually recognised at least one
    // speaker. Without that anchor its chunks would be indistinguishable from
    // surrounding UI chrome, and the selector scraper is a safer fallback.
    if (patternTurns.some((t) => t.speaker)) return patternTurns;

    // Fallback to the selector scraper for surfaces where the pattern scraper
    // produces nothing (e.g. closed-captions panels with known data-tid hooks
    // and very short visible windows that don't expose the name+time header).
    return scrapeTurnsBySelector(container);
  }

  // The original selector/attribute-driven scraper. Good for Teams native
  // surfaces where data-tid attributes are stable.
  function scrapeTurnsBySelector(container) {
    const rowSelectors = [
      // SharePoint Stream (stream.aspx) transcript entries.
      '[data-automationid="transcript-item"]',
      '[data-automationid="transcriptItem"]',
      '[data-automationid="transcriptListItem"]',
      '[data-automationid="transcript-entry"]',
      // Teams native.
      '[data-tid="transcript-list-item"]',
      '[data-tid="transcriptListItem"]',
      '[data-tid="closed-caption-row"]',
      // Generic.
      'div[role="listitem"]',
      ".ts-transcript-entry",
      ".transcript-list-item",
      ".caption-line",
    ];
    const rows = querySelectorAllWithFallbacks(
      container,
      rowSelectors,
      "transcript row"
    );

    const turns = [];
    rows.forEach((row) => {
      const turn = extractTurnFromRow(row);
      if (turn && (turn.text || turn.speaker)) {
        turns.push(turn);
      }
    });
    return turns;
  }

  // Pattern-based scraper. Walks visible text-leaf elements inside the
  // transcript container in document order, then groups them into turns
  // based on the Name -> HH:MM -> Text pattern that every Teams/Stream
  // transcript follows regardless of class naming.
  function scrapeTurnsByPattern(container) {
    const chunks = collectTextChunks(container);
    return groupChunksIntoTurns(chunks);
  }

  // Pattern that matches a bare digital timestamp like "7:24", "12:05", or
  // "1:02:35".
  const DIGITAL_TIME_RE = /\d{1,3}:\d{2}(?::\d{2})?/;

  // Readable timestamp produced by Teams for accessibility, e.g.
  // "0 minutes 3 seconds" or "1 hour 20 minutes 3 seconds".
  const READABLE_TIME_RE =
    /(?:\d+\s*hours?\s+)?\d+\s*minutes?\s+\d+\s*seconds?/i;

  // A chunk consisting solely of a timestamp in digital form, readable form,
  // or the two concatenated together (e.g. "0 minutes 3 seconds0:03").
  const TIMESTAMP_ONLY_RE = new RegExp(
    "^\\s*(?:" +
      "\\d{1,3}:\\d{2}(?::\\d{2})?" +
      "|" +
      "(?:\\d+\\s*hours?\\s+)?\\d+\\s*minutes?\\s+\\d+\\s*seconds?" +
      "|" +
      "\\d{1,3}:\\d{2}(?::\\d{2})?\\s*(?:\\d+\\s*hours?\\s+)?\\d+\\s*minutes?\\s+\\d+\\s*seconds?" +
      "|" +
      "(?:\\d+\\s*hours?\\s+)?\\d+\\s*minutes?\\s+\\d+\\s*seconds?\\s*\\d{1,3}:\\d{2}(?::\\d{2})?" +
      ")\\s*$",
    "i"
  );

  // A "Name TIME ..." trailing-timestamp header. The timestamp may be digital,
  // readable, or both concatenated/adjacent.
  const NAME_PLUS_TIME_RE = new RegExp(
    "^(.{1,80}?)\\s+(" +
      "(?:\\d+\\s*hours?\\s+)?\\d+\\s*minutes?\\s+\\d+\\s*seconds?\\s*(?:\\d{1,3}:\\d{2}(?::\\d{2})?)?" +
      "|" +
      "\\d{1,3}:\\d{2}(?::\\d{2})?\\s*(?:(?:\\d+\\s*hours?\\s+)?\\d+\\s*minutes?\\s+\\d+\\s*seconds?)?" +
      ")\\s*(.*)$",
    "i"
  );

  // System notices Teams inserts into the transcript ("Jacqueline Derron
  // started transcription", "Recording started", etc.). We drop these.
  const SYSTEM_BANNER_RE =
    /^(?:.{1,80}?\s+)?(started|stopped|paused|resumed|joined|left)\s+(transcription|the\s+meeting|the\s+call|recording|captions?)\.?\s*$/i;
  const SYSTEM_BANNER_LEADING_RE =
    /^(transcription|recording|captions?)\s+(started|stopped|paused|resumed)\.?\s*$/i;

  // Strict "looks like a person's name" check used when a chunk stands alone
  // without a timestamp next to it. Requires each word to be capitalized and
  // forbids sentence punctuation or digits, so it won't swallow utterance
  // fragments like "Yeah" or "The timing on.".
  function isStrictName(s) {
    if (!s) return false;
    const trimmed = s.trim();
    if (trimmed.length < 2 || trimmed.length > 60) return false;
    if (/\d/.test(trimmed)) return false;
    if (/[.!?,;:]/.test(trimmed)) return false;
    const words = trimmed.split(/\s+/);
    if (words.length < 1 || words.length > 5) return false;
    for (const w of words) {
      if (!/^[A-Z][A-Za-z'\u2019\-]*\.?$/.test(w)) return false;
    }
    return true;
  }

  function isSystemBanner(text) {
    const t = text.trim();
    return SYSTEM_BANNER_RE.test(t) || SYSTEM_BANNER_LEADING_RE.test(t);
  }

  function isTimestampChunk(text) {
    return TIMESTAMP_ONLY_RE.test(text.trim());
  }

  // Pull a digital timestamp (MM:SS / HH:MM:SS) out of a chunk. If only a
  // readable form is present ("0 minutes 3 seconds") convert it. Returns ""
  // if nothing timestamp-like is present.
  function extractTimestampFromChunk(text) {
    const digital = text.match(DIGITAL_TIME_RE);
    if (digital) return digital[0];

    const readable = text.match(
      /(?:(\d+)\s*hours?\s+)?(\d+)\s*minutes?\s+(\d+)\s*seconds?/i
    );
    if (readable) {
      const h = readable[1] ? parseInt(readable[1], 10) : 0;
      const mm = parseInt(readable[2], 10);
      const ss = parseInt(readable[3], 10);
      const pad = (n) => String(n).padStart(2, "0");
      if (h > 0) return pad(h) + ":" + pad(mm) + ":" + pad(ss);
      return pad(mm) + ":" + pad(ss);
    }
    return "";
  }

  // Try to read a speaker-header chunk. Returns { speaker, timestamp,
  // inlineText } or null. Handles:
  //   - "Name"                                     (bare name, no time)
  //   - "Name 7:24"                                 (digital trailing time)
  //   - "Name 0 minutes 3 seconds"                  (readable trailing time)
  //   - "Name 0 minutes 3 seconds0:03"              (both concatenated)
  //   - "Name 7:24 Text..."                         (time + inline utterance)
  function parseSpeakerHeader(text) {
    const t = text.trim();
    if (!t) return null;

    const m = t.match(NAME_PLUS_TIME_RE);
    if (m && isStrictName(m[1])) {
      const speaker = m[1].trim();
      const timestamp = extractTimestampFromChunk(m[2] || "");
      const inlineText = (m[3] || "").trim();
      return { speaker, timestamp, inlineText };
    }

    if (isStrictName(t)) {
      return { speaker: t, timestamp: "", inlineText: "" };
    }

    return null;
  }

  // Walk the chunk stream in document order, track the currently-speaking
  // person, and emit one turn per utterance chunk. Speaker headers (name
  // alone, or name+timestamp) update the current speaker but don't produce
  // a turn by themselves. Standalone timestamps attach to the next
  // utterance. System banners ("X started transcription") are dropped.
  function groupChunksIntoTurns(chunks) {
    const turns = [];
    let currentSpeaker = "";
    let pendingTimestamp = "";

    for (let i = 0; i < chunks.length; i++) {
      const text = (chunks[i].text || "").trim();
      if (!text) continue;
      if (isSystemBanner(text)) continue;

      const header = parseSpeakerHeader(text);
      if (header) {
        currentSpeaker = header.speaker;
        if (header.timestamp) pendingTimestamp = header.timestamp;
        if (header.inlineText) {
          turns.push({
            speaker: currentSpeaker,
            timestamp: pendingTimestamp,
            text: header.inlineText,
          });
          pendingTimestamp = "";
        }
        continue;
      }

      if (isTimestampChunk(text)) {
        const ts = extractTimestampFromChunk(text);
        if (ts) pendingTimestamp = ts;
        continue;
      }

      turns.push({
        speaker: currentSpeaker,
        timestamp: pendingTimestamp,
        text,
      });
      pendingTimestamp = "";
    }

    return turns;
  }

  // Walk the container's descendants in document order and collect every
  // visible text-leaf element's trimmed text as a chunk. A "text leaf" is an
  // element whose own text content isn't already covered by a child element
  // (so we don't double-count the same words).
  function collectTextChunks(container) {
    const chunks = [];
    const skipTags = new Set([
      "IMG",
      "SVG",
      "BUTTON",
      "STYLE",
      "SCRIPT",
      "VIDEO",
      "AUDIO",
      "CANVAS",
      "INPUT",
      "TEXTAREA",
    ]);
    const walker = document.createTreeWalker(
      container,
      NodeFilter.SHOW_ELEMENT,
      {
        acceptNode(el) {
          if (skipTags.has(el.tagName)) return NodeFilter.FILTER_REJECT;
          if (el.getAttribute && el.getAttribute("aria-hidden") === "true") {
            return NodeFilter.FILTER_REJECT;
          }
          const role = el.getAttribute && el.getAttribute("role");
          if (role === "img" || role === "presentation") {
            return NodeFilter.FILTER_REJECT;
          }
          return NodeFilter.FILTER_ACCEPT;
        },
      }
    );

    let el;
    while ((el = walker.nextNode())) {
      // Only keep text leaves: elements whose direct text nodes have
      // content, and whose element children (if any) don't themselves carry
      // non-trivial text. This gives us one chunk per rendered text span.
      const ownText = getDirectText(el);
      if (!ownText) continue;

      const hasTextChildren = hasAnyChildElementWithText(el);
      if (hasTextChildren) continue;

      chunks.push({ text: ownText, el });
    }
    return chunks;
  }

  // Return the concatenation of the element's direct text-node children,
  // trimmed and whitespace-collapsed. Excludes text from element children.
  function getDirectText(el) {
    let s = "";
    for (const node of el.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        s += node.nodeValue || "";
      }
    }
    return s.replace(/\s+/g, " ").trim();
  }

  // True if any child element has meaningful text content of its own. Used to
  // decide whether an element is a "text leaf" worth capturing as a chunk.
  function hasAnyChildElementWithText(el) {
    for (const child of el.children) {
      if (!child) continue;
      if (child.tagName === "IMG" || child.tagName === "SVG") continue;
      const t = (child.textContent || "").trim();
      if (t.length > 0) return true;
    }
    return false;
  }

  // Pull speaker / timestamp / text from a single transcript row element.
  // Falls back to scanning the row's own text content if no inner selectors
  // match (which makes us resilient to renames).
  function extractTurnFromRow(row) {
    const speakerSelectors = [
      // SharePoint Stream.
      '[data-automationid="transcript-item-author"]',
      '[data-automationid="transcriptAuthor"]',
      '[data-automationid="author"]',
      // Teams native.
      '[data-tid="author"]',
      '[data-tid="transcript-author"]',
      '[data-tid="closed-caption-author"]',
      // Generic.
      ".ts-transcript-author",
      ".author",
      ".speaker-name",
    ];
    const timestampSelectors = [
      // SharePoint Stream.
      '[data-automationid="transcript-item-timestamp"]',
      '[data-automationid="transcriptTimestamp"]',
      '[data-automationid="timestamp"]',
      // Teams native.
      '[data-tid="timestamp"]',
      '[data-tid="transcript-timestamp"]',
      '[data-tid="closed-caption-timestamp"]',
      // Generic.
      ".ts-transcript-timestamp",
      ".timestamp",
      "time",
    ];
    const textSelectors = [
      // SharePoint Stream.
      '[data-automationid="transcript-item-text"]',
      '[data-automationid="transcriptText"]',
      '[data-automationid="transcript-text"]',
      // Teams native.
      '[data-tid="transcript-text"]',
      '[data-tid="closed-caption-text"]',
      // Generic.
      ".ts-transcript-text",
      ".transcript-text",
      ".caption-text",
    ];

    const speakerNode = querySelectorWithFallbacks(
      row,
      speakerSelectors,
      "speaker"
    );
    const timestampNode = querySelectorWithFallbacks(
      row,
      timestampSelectors,
      "timestamp"
    );
    const textNode = querySelectorWithFallbacks(row, textSelectors, "text");

    let speaker = speakerNode ? extractText(speakerNode) : "";
    let timestamp = timestampNode ? extractText(timestampNode) : "";
    let text = textNode ? extractText(textNode) : "";

    // If we couldn't find inner nodes, treat the whole row as a single
    // utterance and try to split out a leading "Name 00:01:23" header.
    if (!text) {
      const rowText = extractText(row);
      const headerMatch = rowText.match(
        /^([^\d\n]{1,80}?)\s+(\d{1,2}:\d{2}(?::\d{2})?)\s+([\s\S]+)$/
      );
      if (headerMatch) {
        speaker = speaker || headerMatch[1].trim();
        timestamp = timestamp || headerMatch[2].trim();
        text = headerMatch[3].trim();
      } else {
        text = rowText;
      }
    }

    return {
      speaker: speaker.trim(),
      timestamp: normalizeTimestamp(timestamp.trim()),
      text: text.trim(),
    };
  }

  // Walk a node and return its text content while skipping things that we
  // know are visual chrome rather than spoken words: avatar images, buttons,
  // emoji reaction layers, screen-reader-only labels, hidden subtrees, etc.
  function extractText(node) {
    if (!node) return "";
    const skipTags = new Set([
      "IMG",
      "SVG",
      "BUTTON",
      "STYLE",
      "SCRIPT",
      "VIDEO",
      "AUDIO",
      "CANVAS",
    ]);
    const parts = [];
    const walker = document.createTreeWalker(
      node,
      NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
      {
        acceptNode(n) {
          if (n.nodeType === Node.ELEMENT_NODE) {
            if (skipTags.has(n.tagName)) {
              return NodeFilter.FILTER_REJECT;
            }
            const aria = n.getAttribute && n.getAttribute("aria-hidden");
            if (aria === "true") {
              return NodeFilter.FILTER_REJECT;
            }
            const role = n.getAttribute && n.getAttribute("role");
            if (role === "img" || role === "presentation") {
              return NodeFilter.FILTER_REJECT;
            }
            return NodeFilter.FILTER_SKIP;
          }
          return NodeFilter.FILTER_ACCEPT;
        },
      }
    );

    let cur = walker.nextNode();
    while (cur) {
      const t = cur.nodeValue;
      if (t) parts.push(t);
      cur = walker.nextNode();
    }
    return parts.join(" ");
  }

  // ---------------------------------------------------------------------------
  // Cleaning
  // ---------------------------------------------------------------------------

  // Clean a single { speaker, timestamp, text } record. The cleaning rules
  // are intentionally minimal so that we never alter the actual words spoken.
  function cleanTurn(turn) {
    const cleanedText = cleanText(turn.text, turn.speaker);
    return {
      speaker: cleanSpeakerName(turn.speaker),
      timestamp: normalizeTimestamp(turn.timestamp || ""),
      text: cleanedText,
    };
  }

  // Trim whitespace and strip trailing punctuation/colons from a name like
  // "Alice B.:" -> "Alice B.".
  function cleanSpeakerName(name) {
    if (!name) return "";
    return name
      .replace(/\s+/g, " ")
      .replace(/[:\u2013\u2014]+\s*$/, "")
      .trim();
  }

  // The core text cleanup pipeline. Applies the rules from the spec in order.
  function cleanText(text, speakerName) {
    if (!text) return "";

    // Strip any HTML tags that may have leaked through (defensive — we use a
    // text walker, but Teams sometimes embeds button labels into text nodes).
    let out = text.replace(/<[^>]+>/g, " ");

    // Smart quotes -> straight quotes.
    out = out
      .replace(/[\u2018\u2019\u201A\u201B]/g, "'")
      .replace(/[\u201C\u201D\u201E\u201F]/g, '"');

    // Em / en dashes inserted by Teams -> regular hyphen.
    out = out.replace(/[\u2013\u2014]/g, "-");

    // Non-breaking spaces and zero-width characters -> regular spaces.
    out = out.replace(/[\u00A0\u2007\u202F]/g, " ");
    out = out.replace(/[\u200B-\u200D\uFEFF]/g, "");

    // Replacement characters become a visible flag so a human can spot them.
    out = out.replace(/\uFFFD/g, "[?]");

    // Collapse repeated whitespace and stray newlines into single spaces.
    out = out.replace(/\s+/g, " ");

    // Remove repeated speaker name injections from the middle of an utterance.
    // Teams sometimes prepends "Name: Name: actual words..." or sprinkles the
    // name between sentences. We strip duplicates of the form "Name:" or "Name "
    // appearing after the first character.
    if (speakerName) {
      const escaped = escapeRegExp(speakerName.replace(/[:.\s]+$/, ""));
      if (escaped) {
        const dupePattern = new RegExp(
          "(?:^|\\s)" + escaped + "\\s*:\\s*",
          "gi"
        );
        out = out.replace(dupePattern, " ");
      }
    }

    // Trim and final whitespace cleanup.
    out = out.replace(/\s+/g, " ").trim();
    return out;
  }

  // Escape a string for use inside a RegExp.
  function escapeRegExp(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  // Normalize timestamps to MM:SS or HH:MM:SS form, stripping wrapping
  // brackets/parens that Teams sometimes uses.
  function normalizeTimestamp(stamp) {
    if (!stamp) return "";
    let t = stamp.trim().replace(/^[\[(]+|[\])]+$/g, "");
    const match = t.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (!match) return "";
    const hh = match[3] !== undefined ? match[1].padStart(2, "0") : null;
    const mm =
      match[3] !== undefined ? match[2].padStart(2, "0") : match[1].padStart(2, "0");
    const ss = (match[3] !== undefined ? match[3] : match[2]).padStart(2, "0");
    return hh ? hh + ":" + mm + ":" + ss : mm + ":" + ss;
  }

  // ---------------------------------------------------------------------------
  // Merging
  // ---------------------------------------------------------------------------

  // Attribute orphan (speakerless) utterances to the most recent known
  // speaker, but keep each utterance as its own turn so the final Markdown
  // renders one utterance per line under each speaker heading instead of
  // collapsing a whole speaker block into one paragraph.
  //
  // Leading speakerless turns (before any speaker has been identified) are
  // dropped as noise when at least one speakered turn exists. If the scraper
  // never identified any speaker, every turn is preserved as-is so we still
  // emit the spoken content.
  function mergeAdjacentTurns(turns) {
    const hasAnySpeaker = turns.some((t) => t.speaker);

    const out = [];
    let lastSpeaker = "";
    for (const turn of turns) {
      if (!turn.text) continue;

      let speaker = turn.speaker;
      if (!speaker && hasAnySpeaker) {
        if (!lastSpeaker) {
          // Noise before the first identified speaker — drop.
          continue;
        }
        speaker = lastSpeaker;
      }

      out.push({
        speaker: speaker || "",
        timestamp: turn.timestamp || "",
        text: turn.text,
      });

      if (turn.speaker) lastSpeaker = turn.speaker;
    }
    return out;
  }

  // Build an alphabetical list of unique speaker names appearing in the
  // transcript.
  function collectParticipants(turns) {
    const seen = new Set();
    for (const t of turns) {
      if (t.speaker) seen.add(t.speaker);
    }
    return Array.from(seen).sort((a, b) => a.localeCompare(b));
  }

  // ---------------------------------------------------------------------------
  // Metadata extraction
  // ---------------------------------------------------------------------------

  // Pull { title, date, url } describing the meeting from the page.
  //
  // The page shows the meeting title and date below the video player in the
  // SharePoint Stream view. We look for a heading element with the title and
  // a nearby date string like "April 2, 2026". Both have sensible fallbacks.
  function extractMeetingMetadata() {
    return {
      title: extractMeetingTitle(),
      date: extractMeetingDate(),
      url: location.href,
    };
  }

  // Find the meeting title. Tries Stream-specific heading hooks first, then
  // any H1/H2 on the page, then falls back to parsing document.title.
  function extractMeetingTitle() {
    const titleSelectors = [
      '[data-automationid="videoTitle"]',
      '[data-automationid="video-title"]',
      '[data-automationid="pageTitle"]',
      '[data-automationid="DocumentTitle"]',
      '[data-automationid="titleField"]',
      "main h1",
      "h1",
      "h2",
    ];
    for (const sel of titleSelectors) {
      const el = document.querySelector(sel);
      if (!el) continue;
      const t = extractText(el).replace(/\s+/g, " ").trim();
      if (t && t.length >= 2 && t.length < 200) {
        console.log("[MeetMark] title matched selector: " + sel);
        return stripTitleNoise(t);
      }
    }
    return stripTitleNoise((document.title || "").trim());
  }

  // Strip filename-looking suffixes and tab-title chrome from a title string.
  // Examples:
  //   "Novartis Spend Classifications-20260402_120001-Meeting Recording.mp4"
  //     -> "Novartis Spend Classifications"
  //   "Novartis Spend Classifications | Microsoft Stream"
  //     -> "Novartis Spend Classifications"
  function stripTitleNoise(title) {
    if (!title) return "";
    let out = title;
    // Drop trailing app/site tags after " | " or " - ".
    out = out.replace(
      /\s*[|\-]\s*(Microsoft Teams|Microsoft Stream|SharePoint|OneDrive).*$/i,
      ""
    );
    // Drop trailing file extension like .mp4 / .webm.
    out = out.replace(/\.[a-z0-9]{2,5}$/i, "");
    // Drop a trailing "-Meeting Recording" tag that Stream appends.
    out = out.replace(/[-_\s]+Meeting Recording\s*$/i, "");
    // Drop a trailing "-YYYYMMDD_HHMMSS" timestamp suffix, with or without
    // further text after it.
    out = out.replace(/[-_\s]+\d{8}[_-]?\d{0,6}.*$/i, "");
    // Collapse internal whitespace.
    out = out.replace(/\s+/g, " ").trim();
    return out;
  }

  // Find the meeting date. Looks for a "Month D, YYYY" string visible on the
  // page near the title; falls back to any such string in the document; falls
  // back to today.
  function extractMeetingDate() {
    const monthName =
      "(January|February|March|April|May|June|July|August|September|October|November|December)";
    const longPattern = new RegExp(monthName + "\\s+(\\d{1,2}),?\\s+(\\d{4})", "i");
    const shortPattern = /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\b/;

    const scopes = [
      document.querySelector("main"),
      document.body,
    ].filter(Boolean);

    for (const scope of scopes) {
      const text = (scope.innerText || "").slice(0, 20000);
      const longMatch = text.match(longPattern);
      if (longMatch) {
        const parsed = new Date(longMatch[0]);
        if (!isNaN(parsed.getTime())) {
          console.log("[MeetMark] date matched long pattern: " + longMatch[0]);
          return parsed;
        }
      }
      const shortMatch = text.match(shortPattern);
      if (shortMatch) {
        const parsed = new Date(
          parseInt(shortMatch[3], 10),
          parseInt(shortMatch[1], 10) - 1,
          parseInt(shortMatch[2], 10)
        );
        if (!isNaN(parsed.getTime())) {
          console.log("[MeetMark] date matched short pattern: " + shortMatch[0]);
          return parsed;
        }
      }
    }

    console.log("[MeetMark] no meeting date found, using today");
    return new Date();
  }

  // ---------------------------------------------------------------------------
  // Formatting
  // ---------------------------------------------------------------------------

  // Render the final Markdown document. Two trailing spaces after each header
  // line force Markdown line breaks.
  function formatMarkdown(turns, participants, metadata) {
    const exportStamp = nowLocalStamp();
    const meetingDateIso = formatDateIso(metadata.date);

    const lines = [];
    lines.push("# " + (metadata.title || "Meeting Transcript"));
    lines.push("");
    if (metadata.title) {
      lines.push("**Title:** " + metadata.title + "  ");
    }
    lines.push("**Date:** " + meetingDateIso + "  ");
    lines.push("**Source:** Microsoft Teams (via SharePoint Stream)  ");
    if (metadata.url) {
      lines.push("**Link:** " + metadata.url + "  ");
    }
    lines.push("");
    lines.push("---");
    lines.push("");
    lines.push("## Transcript");
    lines.push("");

    // Emit one utterance per line, grouped under the speaker's name. The
    // speaker header is only printed when the speaker changes, so a run of
    // utterances from the same person appears as consecutive lines beneath a
    // single name. A blank line separates each speaker group.
    let prevSpeaker = null;
    for (const turn of turns) {
      if (!turn.text) continue;
      const speakerLabel = turn.speaker || "";
      if (speakerLabel !== prevSpeaker) {
        if (prevSpeaker !== null) {
          lines.push("");
        }
        if (speakerLabel) {
          lines.push(speakerLabel);
        }
        prevSpeaker = speakerLabel;
      }
      lines.push(turn.text);
    }
    lines.push("");

    lines.push("---");
    lines.push("*Exported by MeetMark on " + exportStamp + "*");
    lines.push("");

    return lines.join("\n");
  }

  // Format a Date in YYYY-MM-DD form using local time.
  function formatDateIso(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  // Format a Date in YYYY.MM.DD form for use inside a filename.
  function formatDateDotted(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return y + "." + m + "." + day;
  }

  // A human-readable local timestamp for the export footer, e.g.
  // "2026-04-08 at 14:32 PDT".
  function nowLocalStamp() {
    const d = new Date();
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    const tzMatch = d.toString().match(/\(([^)]+)\)$/);
    const tz = tzMatch ? tzMatch[1] : "local time";
    return formatDateIso(d) + " at " + hh + ":" + mm + " " + tz;
  }

  // ---------------------------------------------------------------------------
  // Filename
  // ---------------------------------------------------------------------------

  // Build a download filename in the format
  //   "YYYY.MM.DD - {Meeting Title} Transcript.md"
  // using the title and date extracted from the page, with safe fallbacks.
  function buildFilename(metadata) {
    const datePart = formatDateDotted(metadata.date || new Date());
    const titlePart = sanitizeFilename(metadata.title || "");
    if (titlePart) {
      return datePart + " - " + titlePart + " Transcript.md";
    }
    return datePart + " - Transcript.md";
  }

  // Replace characters that aren't safe in filenames on common OSes.
  function sanitizeFilename(name) {
    if (!name) return "";
    const cleaned = name
      .replace(/[\\/:*?"<>|]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 120);
    return cleaned;
  }
})();

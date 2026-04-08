// MeetMark content script.
// Injected into a Microsoft Teams transcript page on demand. Responsible for:
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
  // chunks; too fast and we skip rendering.
  const SCROLL_STEP_DELAY_MS = 250;

  // After we hit the bottom we wait this long and re-check whether the page
  // grew. If it didn't, we proceed with what we have.
  const SETTLE_DELAY_MS = 800;

  // Listen for the "export" message from the popup. This is the only entry
  // point into the script. We respond asynchronously, so we return true.
  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || message.type !== "MEETMARK_EXPORT") {
      return false;
    }
    runExport()
      .then((result) => sendResponse(result))
      .catch((err) => {
        sendResponse({
          ok: false,
          error: (err && err.message) || String(err),
        });
      });
    return true;
  });

  // Top-level orchestration. Returns { ok, markdown, filename, warning } or
  // { ok: false, error }.
  async function runExport() {
    const container = findTranscriptContainer();
    if (!container) {
      return {
        ok: false,
        error:
          "Could not find transcript content. Make sure you're on a Teams transcript page and it has fully loaded.",
      };
    }

    let warning = null;
    try {
      const scrollResult = await scrollAllContent(container);
      if (scrollResult && scrollResult.timedOut) {
        warning =
          "Scroll timed out before all content loaded. Exporting what was visible.";
      }
    } catch (err) {
      warning = "Scroll error: " + ((err && err.message) || String(err));
    }

    const turns = scrapeTurns(container);
    if (!turns || turns.length === 0) {
      return {
        ok: false,
        error:
          "Could not find transcript content. Make sure you're on a Teams transcript page and it has fully loaded.",
      };
    }

    const cleanedTurns = turns.map(cleanTurn).filter((t) => t.text.length > 0);
    const mergedTurns = mergeAdjacentTurns(cleanedTurns);
    const participants = collectParticipants(mergedTurns);
    const markdown = formatMarkdown(mergedTurns, participants);
    const filename = buildFilename();

    return {
      ok: true,
      markdown,
      filename,
      warning,
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

  // Locate the scroll-container that holds the transcript. Teams reuses a
  // handful of class names depending on the surface, so we try several.
  function findTranscriptContainer() {
    const containerSelectors = [
      '[data-tid="transcript-list"]',
      '[data-tid="transcriptList"]',
      '[data-tid="closed-caption-renderer"]',
      'div[role="list"][aria-label*="ranscript" i]',
      'div[aria-label*="ranscript" i]',
      ".ts-transcript",
      ".transcript-list",
      ".transcript",
    ];
    return querySelectorWithFallbacks(
      document,
      containerSelectors,
      "transcript container"
    );
  }

  // ---------------------------------------------------------------------------
  // Scrolling
  // ---------------------------------------------------------------------------

  // Promise-based sleep helper.
  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  // Determine which element actually scrolls. Some Teams surfaces scroll the
  // container itself; others scroll a parent. Walk up until we find one whose
  // scrollHeight exceeds its clientHeight.
  function findScroller(start) {
    let node = start;
    while (node && node !== document.body) {
      if (node.scrollHeight > node.clientHeight + 4) {
        return node;
      }
      node = node.parentElement;
    }
    return document.scrollingElement || document.documentElement;
  }

  // Scroll incrementally to the bottom, then back to the top, allowing
  // virtualized rows to render. Returns { timedOut: boolean }.
  async function scrollAllContent(container) {
    const scroller = findScroller(container);
    const startTime = Date.now();

    // First push the scroller to the very top so any "scroll up to load
    // earlier" behaviors get a chance.
    scroller.scrollTop = 0;
    await sleep(SETTLE_DELAY_MS);

    let lastHeight = -1;
    let stableLoops = 0;

    while (true) {
      if (Date.now() - startTime > MAX_SCROLL_MS) {
        return { timedOut: true };
      }

      const currentHeight = scroller.scrollHeight;
      const viewport = scroller.clientHeight || window.innerHeight || 600;
      const target = scroller.scrollTop + viewport * 0.8;

      scroller.scrollTop = target;
      await sleep(SCROLL_STEP_DELAY_MS);

      // Bottom reached? wait, then check if anything new appeared.
      if (
        scroller.scrollTop + scroller.clientHeight >=
        scroller.scrollHeight - 4
      ) {
        await sleep(SETTLE_DELAY_MS);
        if (scroller.scrollHeight === currentHeight) {
          stableLoops += 1;
          if (stableLoops >= 2) {
            // Scroll back to top so the user isn't left at the bottom.
            scroller.scrollTop = 0;
            return { timedOut: false };
          }
        } else {
          stableLoops = 0;
        }
      }

      // Track height progress so we don't loop forever on a stuck page.
      if (currentHeight === lastHeight) {
        stableLoops += 1;
        if (stableLoops >= 6) {
          scroller.scrollTop = 0;
          return { timedOut: false };
        }
      } else {
        stableLoops = 0;
        lastHeight = currentHeight;
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Scraping
  // ---------------------------------------------------------------------------

  // Walk the container and pull out one record per transcript row. Each
  // record has { speaker, timestamp, text }.
  function scrapeTurns(container) {
    const rowSelectors = [
      '[data-tid="transcript-list-item"]',
      '[data-tid="transcriptListItem"]',
      '[data-tid="closed-caption-row"]',
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

  // Pull speaker / timestamp / text from a single transcript row element.
  // Falls back to scanning the row's own text content if no inner selectors
  // match (which makes us resilient to renames).
  function extractTurnFromRow(row) {
    const speakerSelectors = [
      '[data-tid="author"]',
      '[data-tid="transcript-author"]',
      '[data-tid="closed-caption-author"]',
      ".ts-transcript-author",
      ".author",
      ".speaker-name",
    ];
    const timestampSelectors = [
      '[data-tid="timestamp"]',
      '[data-tid="transcript-timestamp"]',
      '[data-tid="closed-caption-timestamp"]',
      ".ts-transcript-timestamp",
      ".timestamp",
      "time",
    ];
    const textSelectors = [
      '[data-tid="transcript-text"]',
      '[data-tid="closed-caption-text"]',
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
      timestamp: turn.timestamp || "",
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

  // Merge consecutive rows from the same speaker into a single block. The
  // first timestamp wins; subsequent text is appended as additional sentences.
  function mergeAdjacentTurns(turns) {
    const merged = [];
    for (const turn of turns) {
      const last = merged[merged.length - 1];
      if (last && last.speaker && last.speaker === turn.speaker) {
        if (turn.text) {
          last.text = (last.text + " " + turn.text).replace(/\s+/g, " ").trim();
        }
        if (!last.timestamp && turn.timestamp) {
          last.timestamp = turn.timestamp;
        }
      } else {
        merged.push({
          speaker: turn.speaker,
          timestamp: turn.timestamp,
          text: turn.text,
        });
      }
    }
    return merged;
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
  // Formatting
  // ---------------------------------------------------------------------------

  // Render the final Markdown document. Two trailing spaces after the date /
  // source / participants lines force Markdown line breaks.
  function formatMarkdown(turns, participants) {
    const meetingDate = todayDateISO();
    const exportStamp = nowLocalStamp();

    const lines = [];
    lines.push("# Meeting Transcript");
    lines.push("**Date:** " + meetingDate + "  ");
    lines.push("**Source:** Microsoft Teams  ");
    lines.push(
      "**Participants:** " +
        (participants.length > 0 ? participants.join(", ") : "Unknown")
    );
    lines.push("");
    lines.push("---");
    lines.push("");
    lines.push("## Transcript");
    lines.push("");

    for (const turn of turns) {
      const speakerLabel = turn.speaker || "Unknown speaker";
      const stamp = turn.timestamp ? " `" + turn.timestamp + "`" : "";
      lines.push("**" + speakerLabel + "**" + stamp);
      lines.push(turn.text);
      lines.push("");
    }

    lines.push("---");
    lines.push("*Exported by MeetMark on " + exportStamp + "*");
    lines.push("");

    return lines.join("\n");
  }

  // Today's date in YYYY-MM-DD form, in the user's local timezone.
  function todayDateISO() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  // A human-readable local timestamp for the export footer, e.g.
  // "2026-04-08 at 14:32 PDT".
  function nowLocalStamp() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    const tzMatch = d.toString().match(/\(([^)]+)\)$/);
    const tz = tzMatch ? tzMatch[1] : "local time";
    return y + "-" + m + "-" + day + " at " + hh + ":" + mm + " " + tz;
  }

  // ---------------------------------------------------------------------------
  // Filename
  // ---------------------------------------------------------------------------

  // Build a download filename: prefer the meeting/tab title, fall back to a
  // dated default. Always ends in .md and contains no path separators.
  function buildFilename() {
    const date = todayDateISO();
    const rawTitle = (document.title || "").trim();
    const cleanedTitle = sanitizeFilename(stripCommonTitleNoise(rawTitle));
    if (cleanedTitle) {
      return cleanedTitle + ".md";
    }
    return "transcript-" + date + ".md";
  }

  // Strip common Teams tab-title noise like " | Microsoft Teams".
  function stripCommonTitleNoise(title) {
    if (!title) return "";
    return title
      .replace(/\s*[|\-]\s*Microsoft Teams.*$/i, "")
      .replace(/\s*\(\d+\)\s*/g, " ")
      .trim();
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

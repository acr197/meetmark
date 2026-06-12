// MeetMark full-page screenshot content script.
//
// Ported from mrcoles / GoFullPage's page.js, with one important deviation
// from the archived reference: modern SPAs / dashboards (looker-style BI
// tools, Teams / SharePoint chrome, Salesforce, etc.) frequently put
// `overflow: hidden` on <html> and <body> and do the real scrolling inside
// an inner panel. For those pages `window.scrollTo` is a no-op and
// `document.documentElement.scrollHeight` equals the viewport height, so
// the GoFullPage archive code would emit a single arrangement and produce
// a single-viewport PNG.
//
// To match the behaviour of the shipping GoFullPage Chrome Web Store
// extension we detect the real scrollable element and drive it (scrollTop
// / scrollLeft) instead of window.scrollTo when the root isn't the
// scroller. The popup is told which viewport rect on the captured tab
// image corresponds to the scrollable region so it can crop before
// stitching.
//
// Detection pierces open shadow roots. On Web-Component apps (Home
// Assistant, YouTube, many design systems) the scroller is a <div> nested
// inside one or more shadow roots, which a flat document.querySelectorAll
// cannot see — so without shadow piercing we found no scroller and emitted
// a single-viewport PNG. Closed shadow roots and cross-origin iframes are
// still out of reach for a content script.

(function () {
  // Bump this whenever screenshot.js changes so you can confirm in the page
  // console which copy actually ran. The injected script lives for the life of
  // the tab (see the listener-replacement note below).
  var VERSION = "1.6.1";
  console.log("[MeetMark] screenshot.js " + VERSION + " ready (shadow-DOM aware)");

  // Minimum post-scroll pause before capturing. 400 ms gives virtualized grids
  // (e.g. Smartsheet) enough time to update row transforms after a scroll event.
  var CAPTURE_DELAY = 400;

  // Pad the vertical scroll to overlap sticky headers. 200 is the value
  // GoFullPage uses.
  var STICKY_PAD = 200;

  // Replace any listener installed by a previously-injected copy, so editing
  // screenshot.js takes effect on re-injection WITHOUT a full tab reload. The
  // old `__meetmarkShotLoaded` guard made the FIRST-injected copy win for the
  // life of the tab, which silently kept running stale code after an update —
  // exactly the trap that made a shadow-DOM page look "unfixed". After an
  // extension *reload* the stored reference belongs to an invalidated context,
  // so removeListener is a harmless no-op and we just install the fresh one.
  if (window.__meetmarkShotListener) {
    try {
      chrome.runtime.onConnect.removeListener(window.__meetmarkShotListener);
    } catch (_) {}
  }
  window.__meetmarkShotListener = onMeetmarkShotConnect;
  chrome.runtime.onConnect.addListener(onMeetmarkShotConnect);

  function onMeetmarkShotConnect(port) {
    if (port.name !== "meetmark-shot") return;

    var cancelled = false;
    var ackResolver = null;
    var cleanedUp = false;
    var cleanUp = null;

    port.onDisconnect.addListener(function () {
      cancelled = true;
      if (ackResolver) {
        var r = ackResolver;
        ackResolver = null;
        r();
      }
      if (cleanUp && !cleanedUp) {
        try {
          cleanUp();
        } catch (_) {}
        cleanedUp = true;
      }
    });

    port.onMessage.addListener(function (msg) {
      if (!msg) return;
      if (msg.type === "cancel") {
        cancelled = true;
        if (ackResolver) {
          var r = ackResolver;
          ackResolver = null;
          r();
        }
        return;
      }
      if (msg.type === "captured") {
        if (ackResolver) {
          var r2 = ackResolver;
          ackResolver = null;
          r2();
        }
        return;
      }
      if (msg.type !== "start") return;

      run(port).catch(function (err) {
        try {
          port.postMessage({
            type: "error",
            message: (err && err.message) || String(err),
          });
        } catch (_) {}
        if (cleanUp && !cleanedUp) {
          try {
            cleanUp();
          } catch (_) {}
          cleanedUp = true;
        }
      });
    });

    async function run(port) {
      var body = document.body;
      var docEl = document.documentElement;
      var originalBodyOverflowYStyle = body ? body.style.overflowY : "";
      var originalHtmlOverflowStyle = docEl.style.overflow;
      var originalWinX = window.scrollX;
      var originalWinY = window.scrollY;

      // Probe the root first, then scan descendants for a larger scrollable
      // region. Many dashboards keep the root at viewport size and do the
      // real scrolling inside a nested <div>.
      var target = pickScrollTarget();
      var usingWindow = target.mode === "window";

      var originalElX = 0;
      var originalElY = 0;
      if (!usingWindow) {
        originalElX = target.el.scrollLeft;
        originalElY = target.el.scrollTop;
      }

      // Populated before captures begin; referenced by cleanUp via closure.
      var _fixedHidden = [];
      var _scrollbarStyle = null;

      cleanUp = function () {
        try {
          restoreHiddenElements(_fixedHidden);
          removeStyleEl(_scrollbarStyle);
          docEl.style.overflow = originalHtmlOverflowStyle;
          if (body) body.style.overflowY = originalBodyOverflowYStyle;
          window.scrollTo(originalWinX, originalWinY);
          if (!usingWindow && target.el) {
            target.el.scrollLeft = originalElX;
            target.el.scrollTop = originalElY;
          }
        } catch (_) {}
      };

      // Try to make pages with bad scrolling work, e.g. ones with
      // `body { overflow-y: scroll; }` can break `window.scrollTo`. Only
      // safe to do when we're driving the window; if we've selected an
      // inner element, muting the root overflow can collapse layouts and
      // strand sticky headers.
      if (usingWindow && body) body.style.overflowY = "visible";

      var fullWidth;
      var fullHeight;
      var viewportWidth;
      var viewportHeight;

      if (usingWindow) {
        var widths = [
          docEl.clientWidth,
          body ? body.scrollWidth : 0,
          docEl.scrollWidth,
          body ? body.offsetWidth : 0,
          docEl.offsetWidth,
        ];
        var heights = [
          docEl.clientHeight,
          body ? body.scrollHeight : 0,
          docEl.scrollHeight,
          body ? body.offsetHeight : 0,
          docEl.offsetHeight,
        ];
        fullWidth = max(widths);
        fullHeight = max(heights);
        viewportWidth = window.innerWidth;
        viewportHeight = window.innerHeight;
      } else {
        fullWidth = target.el.scrollWidth;
        fullHeight = target.el.scrollHeight;
        viewportWidth = target.el.clientWidth;
        viewportHeight = target.el.clientHeight;
      }

      var scrollPad = STICKY_PAD;
      var yDelta =
        viewportHeight - (viewportHeight > scrollPad ? scrollPad : 0);
      var xDelta = viewportWidth;

      // During zooming, there can be weird off-by-1 issues.
      if (fullWidth <= xDelta + 1) fullWidth = xDelta;
      if (fullHeight <= yDelta + 1) fullHeight = viewportHeight;

      // Disable page-level scrollbars while we capture. Only when we're
      // driving the window; otherwise we leave the root alone so the inner
      // container keeps its own layout.
      if (usingWindow) docEl.style.overflow = "hidden";

      // Hide position:fixed overlays (sticky headers, comment boxes, chat
      // widgets, etc.) so they don't repeat in every viewport tile.
      // Also inject a stylesheet that suppresses scrollbar tracks.
      _fixedHidden = hideFixedElements();
      _scrollbarStyle = injectScrollbarHider();

      // Build the grid of scroll positions. Bottom-up / left-to-right, same
      // order as GoFullPage. Negative y values clamp to 0 at scroll time;
      // we dedupe the duplicates that produces so we don't capture the same
      // rect twice.
      var arrangements = [];
      var yPos = fullHeight - viewportHeight;
      while (yPos > -yDelta) {
        var xPos = 0;
        while (xPos < fullWidth) {
          arrangements.push([Math.max(0, xPos), Math.max(0, yPos)]);
          xPos += xDelta;
        }
        yPos -= yDelta;
      }

      var seen = Object.create(null);
      arrangements = arrangements.filter(function (p) {
        var k = p[0] + "," + p[1];
        if (seen[k]) return false;
        seen[k] = true;
        return true;
      });

      var numArrangements = arrangements.length;

      try {
        port.postMessage({
          type: "progress",
          message:
            "Capturing " +
            numArrangements +
            " viewport" +
            (numArrangements > 1 ? "s" : "") +
            (usingWindow ? "..." : " (inner panel)..."),
        });
      } catch (_) {}

      for (var i = 0; i < arrangements.length; i++) {
        if (cancelled) {
          if (cleanUp && !cleanedUp) {
            try {
              cleanUp();
            } catch (_) {}
            cleanedUp = true;
          }
          return;
        }

        var next = arrangements[i];
        var x = next[0],
          y = next[1];

        await scrollToAndWait(target, x, y);

        if (cancelled) {
          if (cleanUp && !cleanedUp) {
            try {
              cleanUp();
            } catch (_) {}
            cleanedUp = true;
          }
          return;
        }

        // Read back the real post-scroll position. Some containers clamp
        // scrollTop or defer updates (virtualized lists), so trusting our
        // commanded value is risky.
        var actual = readScroll(target);

        // For inner containers, captureVisibleTab captures the whole tab,
        // so we must tell the popup where on that image the container
        // lives (in CSS px) and how big it is.
        var src = null;
        if (!usingWindow) {
          var rect = target.el.getBoundingClientRect();
          src = {
            left: rect.left,
            top: rect.top,
            width: target.el.clientWidth,
            height: target.el.clientHeight,
          };
        }

        var data = {
          type: "capture",
          x: actual.x,
          y: actual.y,
          complete: (i + 1) / numArrangements,
          windowWidth: viewportWidth,
          totalWidth: fullWidth,
          totalHeight: fullHeight,
          devicePixelRatio: window.devicePixelRatio || 1,
          src: src,
        };

        try {
          port.postMessage(data);
        } catch (_) {
          break;
        }

        await waitForAck();
        if (cancelled) {
          if (cleanUp && !cleanedUp) {
            try {
              cleanUp();
            } catch (_) {}
            cleanedUp = true;
          }
          return;
        }
      }

      if (cleanUp && !cleanedUp) {
        try {
          cleanUp();
        } catch (_) {}
        cleanedUp = true;
      }

      try {
        port.postMessage({ type: "done" });
      } catch (_) {}
    }

    function waitForAck() {
      return new Promise(function (resolve) {
        ackResolver = resolve;
      });
    }
  }

  // Pick what to scroll. The page root and every inner element compete in one
  // pool, scored the same way, so we choose the genuinely-biggest scroller
  // instead of hard-coding a "root vs inner" rule. Detection pierces shadow
  // DOM (see walkDeep), which is what makes this work on Home Assistant and
  // other Web-Component apps.
  function pickScrollTarget() {
    var docEl = document.documentElement;
    var body = document.body;
    var vw = window.innerWidth;
    var vh = window.innerHeight;

    var rootScrollH = Math.max(docEl.scrollHeight, body ? body.scrollHeight : 0);
    var rootDist = Math.max(0, rootScrollH - vh);

    var candidates = findScrollCandidates();

    // A candidate qualifies as the main content well only if it scrolls a
    // meaningful amount AND is both wide and tall on screen. The width gate is
    // what rejects scrolling sidebars / nav rails: they are tall but narrow.
    var MIN_SCROLL = 8;
    var eligible = candidates.filter(function (c) {
      return (
        c.scrollDistance >= MIN_SCROLL &&
        c.visWFrac >= 0.4 &&
        c.visHFrac >= 0.35
      );
    });

    logCandidates(candidates, eligible);

    // Nothing qualified. Drive the window if the document scrolls; otherwise
    // the page already fits in one viewport (a single screenshot is correct).
    if (eligible.length === 0) {
      console.log(
        "[MeetMark] no inner scroller qualified; using window (root scrollDistance=" +
          rootDist +
          ")"
      );
      return { mode: "window", el: null };
    }

    // Most vertical scroll distance wins. Among candidates within 20% of the
    // leader, prefer the larger visible area, then the one centered
    // horizontally (content wells sit center; side panels do not).
    var viewportArea = vw * vh || 1;
    var centerX = vw / 2;
    eligible.sort(function (a, b) {
      return b.scrollDistance - a.scrollDistance;
    });
    var lead = eligible[0].scrollDistance;
    var near = eligible.filter(function (c) {
      return c.scrollDistance >= lead * 0.8;
    });
    near.sort(function (a, b) {
      if (Math.abs(b.area - a.area) > viewportArea * 0.05) {
        return b.area - a.area;
      }
      var aMid = Math.abs(a.rect.left + a.rect.width / 2 - centerX);
      var bMid = Math.abs(b.rect.left + b.rect.width / 2 - centerX);
      return aMid - bMid;
    });
    var chosen = near[0];

    // If the window scrolls farther than the best inner candidate, the window
    // is probably the real scroller (e.g. a long article that merely contains
    // a scrollable code block). Prefer it.
    if (rootDist > chosen.scrollDistance) {
      console.log(
        "[MeetMark] window scrolls farther (" +
          rootDist +
          ") than best inner (" +
          chosen.scrollDistance +
          "); using window"
      );
      return { mode: "window", el: null };
    }

    console.log(
      "[MeetMark] scroll target:",
      describeEl(chosen.el),
      "scrollDistance=" + chosen.scrollDistance,
      "areaFrac=" + chosen.areaFrac.toFixed(2)
    );
    if (near.length > 1) {
      console.log(
        "[MeetMark] note: " +
          near.length +
          " comparably-sized scrollers exist; captured the largest/most central one."
      );
    }
    return { mode: "element", el: chosen.el };
  }

  // Visit every element in the document, descending into open shadow roots.
  // `visit(el)` may return false to stop the walk early. Node-capped so a
  // pathological DOM can't hang the capture. This is the key to finding
  // scrollers on Web-Component apps, where the scroller lives inside nested
  // shadow roots that a flat document.querySelectorAll("*") never sees.
  function walkDeep(visit) {
    var MAX_NODES = 120000;
    var seen = 0;
    var stack = [document];
    while (stack.length) {
      var root = stack.pop();
      var kids;
      try {
        kids = root.querySelectorAll("*");
      } catch (_) {
        continue;
      }
      for (var i = 0; i < kids.length; i++) {
        var el = kids[i];
        if (++seen > MAX_NODES) return;
        if (visit(el) === false) return;
        // Open shadow roots only; closed roots report null and are skipped.
        if (el.shadowRoot) stack.push(el.shadowRoot);
      }
    }
  }

  // Collect every element that scrolls vertically, with the geometry we need
  // to decide which one is the page's main content region. Pierces shadow DOM.
  function findScrollCandidates() {
    var vw = window.innerWidth;
    var vh = window.innerHeight;
    var viewportArea = vw * vh || 1;
    var out = [];

    walkDeep(function (el) {
      if (el === document.documentElement || el === document.body) return;

      var vDist = el.scrollHeight - el.clientHeight;
      var hDist = el.scrollWidth - el.clientWidth;
      if (vDist <= 1 && hDist <= 1) return;

      var cs;
      try {
        cs = getComputedStyle(el);
      } catch (_) {
        return;
      }
      var oy = cs.overflowY;
      var ox = cs.overflowX;
      var scrollsY =
        (oy === "auto" || oy === "scroll" || oy === "overlay") && vDist > 1;
      var scrollsX =
        (ox === "auto" || ox === "scroll" || ox === "overlay") && hDist > 1;
      if (!scrollsY && !scrollsX) return;

      var rect;
      try {
        rect = el.getBoundingClientRect();
      } catch (_) {
        return;
      }

      // Portion of the element actually on screen. A real content well fills
      // most of the viewport; popups, tooltips and log panes do not.
      var visW = Math.max(0, Math.min(rect.right, vw) - Math.max(rect.left, 0));
      var visH = Math.max(0, Math.min(rect.bottom, vh) - Math.max(rect.top, 0));
      var area = visW * visH;

      out.push({
        el: el,
        rect: rect,
        area: area,
        areaFrac: area / viewportArea,
        visWFrac: visW / vw,
        visHFrac: visH / vh,
        scrollsY: scrollsY,
        scrollDistance: scrollsY ? vDist : 0,
        hScrollDistance: scrollsX ? hDist : 0,
        scrollHeight: el.scrollHeight,
      });
    });

    return out;
  }

  // Human-readable element descriptor for the console, including which shadow
  // host it lives under (so you can find it again in the Elements panel).
  function describeEl(el) {
    var cls =
      typeof el.className === "string"
        ? el.className.trim()
        : (el.getAttribute && el.getAttribute("class")) || "";
    var host = "";
    try {
      var root = el.getRootNode && el.getRootNode();
      if (root && root.host) {
        host = " (inside <" + root.host.tagName.toLowerCase() + "> shadow root)";
      }
    } catch (_) {}
    return (
      "<" +
      el.tagName.toLowerCase() +
      (el.id ? "#" + el.id : "") +
      (cls ? "." + cls.split(/\s+/).slice(0, 3).join(".") : "") +
      ">" +
      host
    );
  }

  // Dump the full candidate list so a misbehaving page can be diagnosed from
  // the tab's console without rebuilding the extension.
  function logCandidates(all, eligible) {
    try {
      console.log(
        "[MeetMark] scroll scan: " +
          all.length +
          " scrollable element(s) found (shadow DOM included), " +
          eligible.length +
          " qualify as main content"
      );
      all
        .slice()
        .sort(function (a, b) {
          return b.scrollDistance - a.scrollDistance;
        })
        .slice(0, 12)
        .forEach(function (c) {
          console.log(
            "  " +
              (eligible.indexOf(c) >= 0 ? "[main]" : "      ") +
              " " +
              describeEl(c.el) +
              "  scrollDist=" +
              c.scrollDistance +
              " visWFrac=" +
              c.visWFrac.toFixed(2) +
              " visHFrac=" +
              c.visHFrac.toFixed(2)
          );
        });
    } catch (_) {}
  }

  // Scroll the target to (x, y) using incremental scrollTop += delta so that
  // SPAs that ignore absolute assignments still advance. After commanding the
  // scroll, poll until the position stabilises, then wait for two animation
  // frames so virtualized grids (which update row transforms inside rAF
  // callbacks) have completed their repaint before we capture.
  async function scrollToAndWait(target, x, y) {
    if (target.mode === "window") {
      window.scrollTo(x, y);
      await delay(CAPTURE_DELAY);
      return;
    }
    var el = target.el;
    var prevTop = el.scrollTop;
    var deltaY = y - prevTop;
    el.scrollLeft = x;
    el.scrollTop += deltaY;

    await waitForScrollStable(el, CAPTURE_DELAY, 800);
    // Double-rAF: first frame lets the scroll event handler run; second
    // frame lets any follow-up layout / transform update flush to the GPU.
    await waitForRaf();

    // Stall detection: if the element didn't reach the target, retry once.
    var reached = el.scrollTop;
    if (Math.abs(reached - y) > 5) {
      console.log("[MeetMark] scroll stalled at", reached, "(target " + y + "), retrying");
      await delay(800);
      el.scrollTop += y - el.scrollTop;
      await waitForScrollStable(el, CAPTURE_DELAY, 800);
      await waitForRaf();
    }

    console.log("[MeetMark] scrollTop=" + el.scrollTop + " (target " + y + ")");
  }

  // Poll el.scrollTop every 50 ms until it stops changing, waiting at least
  // minMs and giving up after maxMs.
  async function waitForScrollStable(el, minMs, maxMs) {
    var start = Date.now();
    var last = el.scrollTop;
    while (true) {
      await delay(50);
      var elapsed = Date.now() - start;
      var cur = el.scrollTop;
      if (cur === last && elapsed >= minMs) break;
      if (elapsed >= maxMs) break;
      last = cur;
    }
  }

  // Wait for two animation frames so all rAF-driven render work flushes.
  function waitForRaf() {
    return new Promise(function (resolve) {
      requestAnimationFrame(function () {
        requestAnimationFrame(resolve);
      });
    });
  }

  function readScroll(target) {
    if (target.mode === "window") {
      return { x: window.scrollX, y: window.scrollY };
    }
    return { x: target.el.scrollLeft, y: target.el.scrollTop };
  }

  function delay(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  function max(nums) {
    return Math.max.apply(
      Math,
      nums.filter(function (x) {
        return x;
      })
    );
  }

  // Hide position:fixed and position:sticky elements so they don't repeat in
  // every captured tile. Fixed overlays (toolbars, comment boxes, chat widgets)
  // always appear at the same viewport position regardless of scroll, so they'd
  // stamp into every tile. Sticky elements (table headers, JIRA nav bars) behave
  // similarly — they "stick" to the viewport edge in every capture frame.
  // We set visibility:hidden on both kinds; layout is preserved (no reflow) while
  // the elements are invisible in the captures. Returns a restore list.
  function hideFixedElements() {
    var hidden = [];
    try {
      // Pierce shadow DOM so sticky headers inside component apps (e.g. Home
      // Assistant's top app bar) are hidden too, instead of stamping into
      // every tile.
      walkDeep(function (el) {
        if (el === document.documentElement || el === document.body) return;
        var cs;
        try {
          cs = getComputedStyle(el);
        } catch (_) {
          return;
        }
        var pos = cs.position;
        if (pos !== "fixed" && pos !== "sticky") return;
        // Skip elements that are already invisible — no need to track them.
        if (cs.display === "none" || cs.visibility === "hidden") return;
        // Skip zero-size elements (they wouldn't appear in the capture anyway).
        try {
          var r = el.getBoundingClientRect();
          if (r.width === 0 && r.height === 0) return;
        } catch (_) {}
        hidden.push({ el: el, visibility: el.style.visibility });
        el.style.setProperty("visibility", "hidden", "important");
      });
    } catch (_) {}
    return hidden;
  }

  function restoreHiddenElements(hidden) {
    if (!hidden) return;
    for (var i = 0; i < hidden.length; i++) {
      try {
        hidden[i].el.style.visibility = hidden[i].visibility;
      } catch (_) {}
    }
  }

  // Inject a <style> that hides scrollbar tracks so they don't appear as
  // repeated artefacts along the edges of captured tiles.
  function injectScrollbarHider() {
    try {
      var style = document.createElement("style");
      style.id = "__meetmark_noscrollbar__";
      style.textContent =
        "::-webkit-scrollbar{display:none!important}" +
        "*{scrollbar-width:none!important}";
      (document.head || document.documentElement).appendChild(style);
      return style;
    } catch (_) {
      return null;
    }
  }

  function removeStyleEl(el) {
    if (el && el.parentNode) {
      try {
        el.parentNode.removeChild(el);
      } catch (_) {}
    }
  }
})();

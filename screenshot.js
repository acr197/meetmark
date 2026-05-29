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

(function () {
  if (window.__meetmarkShotLoaded) {
    return;
  }
  window.__meetmarkShotLoaded = true;

  // Minimum post-scroll pause before capturing. 400 ms gives virtualized grids
  // (e.g. Smartsheet) enough time to update row transforms after a scroll event.
  var CAPTURE_DELAY = 400;

  // Pad the vertical scroll to overlap sticky headers. 200 is the value
  // GoFullPage uses.
  var STICKY_PAD = 200;

  chrome.runtime.onConnect.addListener(function (port) {
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
  });

  // Pick the scrollable target. If an inner element has taller scrollable
  // content than the document root, use it; otherwise drive the window.
  function pickScrollTarget() {
    var docEl = document.documentElement;
    var body = document.body;

    var inner = findTallestInnerScroller();

    var rootScrollHeight = Math.max(
      docEl.scrollHeight,
      body ? body.scrollHeight : 0
    );
    var innerScrollHeight = inner ? inner.scrollHeight : 0;

    var rootScrolls =
      docEl.scrollHeight > docEl.clientHeight + 1 ||
      (body && body.scrollHeight > body.clientHeight + 1);

    // Inner wins when it has meaningfully more hidden content than the root,
    // or when the root doesn't actually scroll (overflow:hidden dashboards).
    if (inner && (!rootScrolls || innerScrollHeight > rootScrollHeight * 1.1)) {
      return { mode: "element", el: inner };
    }
    return { mode: "window", el: null };
  }

  // Walk the DOM and collect all vertically scrollable elements (overflow
  // auto/scroll, scrollHeight > clientHeight). Pick the tallest one; if two
  // candidates are within 20% of each other, prefer the one whose horizontal
  // midpoint is closest to the viewport center (content wells beat sidebars).
  function findTallestInnerScroller() {
    var candidates = [];
    var nodes = document.querySelectorAll("*");
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el === document.documentElement || el === document.body) continue;

      if (el.scrollHeight <= el.clientHeight + 1) continue;

      var cs;
      try {
        cs = getComputedStyle(el);
      } catch (_) {
        continue;
      }
      var oy = cs.overflowY;
      if (oy !== "auto" && oy !== "scroll" && oy !== "overlay") continue;

      var rect;
      try {
        rect = el.getBoundingClientRect();
      } catch (_) {
        continue;
      }
      // Skip elements narrower/shorter than 40% of the viewport — sidebars
      // and tooltip containers are never the main content well.
      if (
        rect.width < window.innerWidth * 0.4 ||
        rect.height < window.innerHeight * 0.4
      ) {
        continue;
      }

      candidates.push({ el: el, scrollHeight: el.scrollHeight, rect: rect });
    }

    if (candidates.length === 0) return null;

    // Sort tallest first.
    candidates.sort(function (a, b) {
      return b.scrollHeight - a.scrollHeight;
    });

    // Among candidates within 20% of the tallest, prefer the one whose
    // midpoint is closest to the horizontal center of the viewport.
    var topHeight = candidates[0].scrollHeight;
    var similar = candidates.filter(function (c) {
      return c.scrollHeight >= topHeight * 0.8;
    });

    var chosen;
    if (similar.length > 1) {
      var centerX = window.innerWidth / 2;
      chosen = similar.reduce(function (a, b) {
        var aMid = a.rect.left + a.rect.width / 2;
        var bMid = b.rect.left + b.rect.width / 2;
        return Math.abs(aMid - centerX) <= Math.abs(bMid - centerX) ? a : b;
      });
    } else {
      chosen = candidates[0];
    }

    var cls =
      typeof chosen.el.className === "string"
        ? chosen.el.className.trim() || "(none)"
        : "(none)";
    console.log(
      "[MeetMark] scroll target selected:",
      chosen.el.tagName,
      "id=" + (chosen.el.id || "(none)"),
      "class=" + cls,
      "scrollHeight=" + chosen.scrollHeight
    );

    return chosen.el;
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
      var nodes = document.querySelectorAll("*");
      for (var i = 0; i < nodes.length; i++) {
        var el = nodes[i];
        if (el === document.documentElement || el === document.body) continue;
        var cs;
        try {
          cs = getComputedStyle(el);
        } catch (_) {
          continue;
        }
        var pos = cs.position;
        if (pos !== "fixed" && pos !== "sticky") continue;
        // Skip elements that are already invisible — no need to track them.
        if (cs.display === "none" || cs.visibility === "hidden") continue;
        // Skip zero-size elements (they wouldn't appear in the capture anyway).
        try {
          var r = el.getBoundingClientRect();
          if (r.width === 0 && r.height === 0) continue;
        } catch (_) {}
        hidden.push({ el: el, visibility: el.style.visibility });
        el.style.setProperty("visibility", "hidden", "important");
      }
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

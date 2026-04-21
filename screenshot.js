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

  // Matches GoFullPage's CAPTURE_DELAY. Chrome throttles
  // chrome.tabs.captureVisibleTab to roughly 2 calls per second — the popup
  // serializes captures with the port ack, so this mainly controls how long
  // we pause after a scroll for Fluent / React / virtualized UIs to lay out.
  var CAPTURE_DELAY = 250;

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

        scrollTo(target, x, y);

        await delay(CAPTURE_DELAY);
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

  // Pick the scrollable target. Prefer the window when the document root
  // actually scrolls; otherwise find the largest scrollable descendant.
  function pickScrollTarget() {
    var docEl = document.documentElement;
    var body = document.body;
    var rootScrolls =
      (docEl.scrollHeight > docEl.clientHeight + 1 ||
        docEl.scrollWidth > docEl.clientWidth + 1) ||
      (body &&
        (body.scrollHeight > body.clientHeight + 1 ||
          body.scrollWidth > body.clientWidth + 1));

    var inner = findBiggestInnerScroller();

    // Choose the one with more hidden content. Inner wins if it dwarfs the
    // window (the dashboard case), window wins for normal pages.
    var rootArea = 0;
    if (rootScrolls) {
      var w = Math.max(docEl.scrollWidth, body ? body.scrollWidth : 0);
      var h = Math.max(docEl.scrollHeight, body ? body.scrollHeight : 0);
      rootArea = w * h;
    }
    var innerArea = inner ? inner.scrollWidth * inner.scrollHeight : 0;

    if (innerArea > rootArea * 1.1 && inner) {
      return { mode: "element", el: inner };
    }
    return { mode: "window", el: null };
  }

  function findBiggestInnerScroller() {
    var best = null;
    var bestArea = 0;
    var nodes = document.querySelectorAll("*");
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el === document.documentElement || el === document.body) continue;

      // Cheap size check first — elements smaller than half the viewport
      // in either axis are almost never the page's main scroller.
      var rect;
      try {
        rect = el.getBoundingClientRect();
      } catch (_) {
        continue;
      }
      if (
        rect.width < window.innerWidth * 0.4 ||
        rect.height < window.innerHeight * 0.4
      ) {
        continue;
      }

      var scrollsY = el.scrollHeight > el.clientHeight + 1;
      var scrollsX = el.scrollWidth > el.clientWidth + 1;
      if (!scrollsY && !scrollsX) continue;

      var cs;
      try {
        cs = getComputedStyle(el);
      } catch (_) {
        continue;
      }
      var oy = cs.overflowY;
      var ox = cs.overflowX;
      var canScroll =
        oy === "auto" ||
        oy === "scroll" ||
        oy === "overlay" ||
        ox === "auto" ||
        ox === "scroll" ||
        ox === "overlay";
      if (!canScroll) continue;

      var area = el.scrollWidth * el.scrollHeight;
      if (area > bestArea) {
        bestArea = area;
        best = el;
      }
    }
    return best;
  }

  function scrollTo(target, x, y) {
    if (target.mode === "window") {
      window.scrollTo(x, y);
    } else {
      target.el.scrollLeft = x;
      target.el.scrollTop = y;
    }
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

  // Find all position:fixed elements (toolbars, comment boxes, overlays) and
  // hide them so they don't appear in every captured tile. Returns a list of
  // {el, visibility} records for restoreHiddenElements().
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
        if (cs.position !== "fixed") continue;
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

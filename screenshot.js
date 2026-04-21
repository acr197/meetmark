// MeetMark full-page screenshot content script.
//
// Ported faithfully from mrcoles / GoFullPage's page.js so the behavior
// matches what the user sees in the GoFullPage Chrome extension. The
// original code is MV2; the only adaptations here are MV3 concerns:
//
//   * Port-based messaging instead of chrome.runtime.sendMessage with a
//     callback (popups in MV3 react better to ports when they send many
//     messages in sequence).
//   * Awaiting the popup's "captured" ack before scrolling to the next
//     arrangement, which keeps chrome.tabs.captureVisibleTab under its
//     rate limit.
//
// The arrangement math, overflow swapping, sticky-header pad, and measure
// logic are copied from the GoFullPage source the user referenced.

(function () {
  if (window.__meetmarkShotLoaded) {
    return;
  }
  window.__meetmarkShotLoaded = true;

  // Matches GoFullPage's CAPTURE_DELAY. Chrome throttles
  // chrome.tabs.captureVisibleTab to roughly 2 calls per second — the popup
  // serializes captures with the port ack, so this only controls how long we
  // pause after a scroll for Fluent / React UIs to lay out.
  var CAPTURE_DELAY = 150;

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
      var originalBodyOverflowYStyle = body ? body.style.overflowY : "";
      var originalX = window.scrollX;
      var originalY = window.scrollY;
      var originalOverflowStyle = document.documentElement.style.overflow;

      cleanUp = function () {
        document.documentElement.style.overflow = originalOverflowStyle;
        if (body) body.style.overflowY = originalBodyOverflowYStyle;
        window.scrollTo(originalX, originalY);
      };

      // Try to make pages with bad scrolling work, e.g., ones with
      // body { overflow-y: scroll; } can break window.scrollTo.
      if (body) body.style.overflowY = "visible";

      var widths = [
        document.documentElement.clientWidth,
        body ? body.scrollWidth : 0,
        document.documentElement.scrollWidth,
        body ? body.offsetWidth : 0,
        document.documentElement.offsetWidth,
      ];
      var heights = [
        document.documentElement.clientHeight,
        body ? body.scrollHeight : 0,
        document.documentElement.scrollHeight,
        body ? body.offsetHeight : 0,
        document.documentElement.offsetHeight,
      ];

      var fullWidth = max(widths);
      var fullHeight = max(heights);
      var windowWidth = window.innerWidth;
      var windowHeight = window.innerHeight;
      var scrollPad = STICKY_PAD;
      var yDelta =
        windowHeight - (windowHeight > scrollPad ? scrollPad : 0);
      var xDelta = windowWidth;

      // During zooming, there can be weird off-by-1 issues.
      if (fullWidth <= xDelta + 1) fullWidth = xDelta;

      // Disable all scrollbars while we capture.
      document.documentElement.style.overflow = "hidden";

      // Build the grid of scroll positions. Bottom-up / left-to-right,
      // same order as GoFullPage.
      var arrangements = [];
      var yPos = fullHeight - windowHeight;
      while (yPos > -yDelta) {
        var xPos = 0;
        while (xPos < fullWidth) {
          arrangements.push([xPos, yPos]);
          xPos += xDelta;
        }
        yPos -= yDelta;
      }

      var numArrangements = arrangements.length;

      try {
        port.postMessage({
          type: "progress",
          message:
            "Capturing " +
            numArrangements +
            " viewport" +
            (numArrangements > 1 ? "s" : "") +
            "...",
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

        window.scrollTo(x, y);

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

        var data = {
          type: "capture",
          x: window.scrollX,
          y: window.scrollY,
          complete: (i + 1) / numArrangements,
          windowWidth: windowWidth,
          totalWidth: fullWidth,
          totalHeight: fullHeight,
          devicePixelRatio: window.devicePixelRatio || 1,
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
})();

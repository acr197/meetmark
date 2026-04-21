// MeetMark full-page screenshot content script.
//
// Injected by the popup when the user clicks "Capture full page as PNG".
// Its job is to drive the scroll loop that lets the popup capture the
// entire page one viewport at a time:
//
//   1. Save the page's current scroll position and overflow styles.
//   2. Measure the full page dimensions and plan the grid of (x, y)
//      positions we need to scroll through. A vertical pad handles sticky
//      headers (we advance by windowHeight - 200px each row) the same way
//      Peter Coles' full-page-screen-capture-chrome-extension does.
//   3. For each position, scrollTo(x, y), wait CAPTURE_DELAY for things to
//      settle, and tell the popup to call chrome.tabs.captureVisibleTab()
//      while we hold the page at that position. The popup replies when the
//      capture is stitched into its canvas.
//   4. Restore scroll + overflow styles and tell the popup we're done.
//
// The popup is the only code path with access to the chrome.tabs APIs we
// need. All we do here is drive the DOM.

(function () {
  if (window.__meetmarkShotLoaded) {
    return;
  }
  window.__meetmarkShotLoaded = true;

  // Delay between scroll and capture. Chrome throttles
  // chrome.tabs.captureVisibleTab to ~2 calls per second; we must leave
  // headroom for that plus layout settling after the scroll.
  const CAPTURE_DELAY = 520;

  // Vertical overlap between consecutive scroll positions. Gives sticky
  // headers somewhere to render over without covering unseen content.
  const STICKY_PAD = 200;

  chrome.runtime.onConnect.addListener((port) => {
    if (port.name !== "meetmark-shot") return;

    let cancelled = false;
    let awaitingAck = null; // function to call when the popup acks a capture
    let cleanedUp = false;
    let restore = null; // closure that undoes scroll + style changes

    port.onDisconnect.addListener(() => {
      cancelled = true;
      if (awaitingAck) {
        const fn = awaitingAck;
        awaitingAck = null;
        fn();
      }
      if (restore && !cleanedUp) {
        try {
          restore();
        } catch (_) {}
        cleanedUp = true;
      }
    });

    port.onMessage.addListener((msg) => {
      if (!msg) return;
      if (msg.type === "cancel") {
        cancelled = true;
        if (awaitingAck) {
          const fn = awaitingAck;
          awaitingAck = null;
          fn();
        }
        return;
      }
      if (msg.type === "captured") {
        if (awaitingAck) {
          const fn = awaitingAck;
          awaitingAck = null;
          fn();
        }
        return;
      }
      if (msg.type !== "start") return;

      run(port)
        .catch((err) => {
          try {
            port.postMessage({
              type: "error",
              message: (err && err.message) || String(err),
            });
          } catch (_) {}
        })
        .finally(() => {
          if (restore && !cleanedUp) {
            try {
              restore();
            } catch (_) {}
            cleanedUp = true;
          }
        });
    });

    async function run(port) {
      const body = document.body;
      const originalBodyOverflowY = body ? body.style.overflowY : "";
      const originalDocOverflow = document.documentElement.style.overflow;
      const originalX = window.scrollX;
      const originalY = window.scrollY;

      restore = () => {
        document.documentElement.style.overflow = originalDocOverflow;
        if (body) body.style.overflowY = originalBodyOverflowY;
        window.scrollTo(originalX, originalY);
      };

      // Some pages set body { overflow-y: scroll } which breaks
      // window.scrollTo. Flip it off for the duration of the capture.
      if (body) body.style.overflowY = "visible";

      const widths = [
        document.documentElement.clientWidth,
        body ? body.scrollWidth : 0,
        document.documentElement.scrollWidth,
        body ? body.offsetWidth : 0,
        document.documentElement.offsetWidth,
      ];
      const heights = [
        document.documentElement.clientHeight,
        body ? body.scrollHeight : 0,
        document.documentElement.scrollHeight,
        body ? body.offsetHeight : 0,
        document.documentElement.offsetHeight,
      ];

      let fullWidth = max(widths);
      let fullHeight = max(heights);
      const windowWidth = window.innerWidth;
      const windowHeight = window.innerHeight;
      const yDelta =
        windowHeight - (windowHeight > STICKY_PAD ? STICKY_PAD : 0);
      const xDelta = windowWidth;

      if (fullWidth <= xDelta + 1) fullWidth = xDelta;

      // Hide scrollbars so they don't appear in the stitched image and
      // don't throw off width measurements.
      document.documentElement.style.overflow = "hidden";

      // Plan the grid. Start from the bottom so a page with lazy-loaded
      // content at the top isn't disturbed before the top row fires.
      const arrangements = [];
      let yPos = fullHeight - windowHeight;
      while (yPos > -yDelta) {
        let xPos = 0;
        while (xPos < fullWidth) {
          arrangements.push([Math.max(0, xPos), Math.max(0, yPos)]);
          xPos += xDelta;
        }
        yPos -= yDelta;
      }
      if (!arrangements.length) {
        // Page is shorter than one viewport — still capture once.
        arrangements.push([0, 0]);
      }

      const total = arrangements.length;
      try {
        port.postMessage({
          type: "progress",
          message:
            "Capturing " + total + " viewport" + (total > 1 ? "s" : "") + "...",
        });
      } catch (_) {}

      for (let i = 0; i < arrangements.length; i++) {
        if (cancelled) return;

        const [x, y] = arrangements[i];
        window.scrollTo(x, y);

        await delay(CAPTURE_DELAY);
        if (cancelled) return;

        // Read the resolved scroll position back — some pages clamp or
        // round the scrollTo value, and the popup stitches against what
        // the page actually scrolled to.
        const actualX = window.scrollX;
        const actualY = window.scrollY;

        try {
          port.postMessage({
            type: "capture",
            x: actualX,
            y: actualY,
            complete: (i + 1) / total,
            windowWidth,
            windowHeight,
            totalWidth: fullWidth,
            totalHeight: fullHeight,
            devicePixelRatio: window.devicePixelRatio || 1,
          });
        } catch (_) {
          return;
        }

        await waitForAck();
        if (cancelled) return;
      }

      try {
        port.postMessage({ type: "done" });
      } catch (_) {}
    }

    function waitForAck() {
      return new Promise((resolve) => {
        awaitingAck = resolve;
      });
    }
  });

  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
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

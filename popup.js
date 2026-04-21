// MeetMark popup controller.
//
// The popup exposes three export paths:
//
//   1. Quick Grab MD            — one-click transcript export as Markdown,
//                                 the historical default behavior.
//   2. Grab as (MD / TXT / PDF) — same transcript pipeline as Quick Grab,
//                                 but the content is re-rendered in the
//                                 chosen format before download. PDF opens
//                                 a printable HTML tab so the user can Save
//                                 as PDF from the browser print dialog.
//   3. Capture full page as PNG — full-page screenshot of the current tab
//                                 by scrolling the page, calling
//                                 chrome.tabs.captureVisibleTab() at each
//                                 arrangement, and stitching the captures
//                                 into one PNG. Inspired by Peter Coles'
//                                 full-page-screen-capture-chrome-extension.
//                                 Works on any tab (relies on activeTab),
//                                 not just Teams.

document.addEventListener("DOMContentLoaded", () => {
  const quickBtn = document.getElementById("quickBtn");
  const grabBtn = document.getElementById("grabBtn");
  const pngBtn = document.getElementById("pngBtn");
  const formatSelect = document.getElementById("formatSelect");
  const cancelBtn = document.getElementById("cancelBtn");
  const status = document.getElementById("status");
  const detail = document.getElementById("detail");
  const progressBar = document.getElementById("progressBar");
  const progressFill = progressBar.querySelector(".fill");

  // Exactly one of these is active at a time. The cancel button talks to
  // whichever is set.
  let activePort = null;
  let activeCancel = null;

  function setStatus(message, level) {
    status.textContent = message || "";
    status.className = "status " + (level || "info");
  }

  function setDetail(text) {
    detail.textContent = text || "";
  }

  function setBusy(busy, determinate) {
    quickBtn.disabled = busy;
    grabBtn.disabled = busy;
    pngBtn.disabled = busy;
    formatSelect.disabled = busy;
    cancelBtn.disabled = !busy;
    if (busy) {
      progressBar.classList.add("active");
      if (determinate) {
        progressBar.classList.add("determinate");
        progressFill.style.width = "0%";
      } else {
        progressBar.classList.remove("determinate");
        progressFill.style.width = "";
      }
    } else {
      progressBar.classList.remove("active");
      progressBar.classList.remove("determinate");
      progressFill.style.width = "";
    }
  }

  function setProgress(fraction) {
    if (typeof fraction !== "number" || !isFinite(fraction)) return;
    const pct = Math.max(0, Math.min(1, fraction)) * 100;
    progressFill.style.width = pct.toFixed(1) + "%";
  }

  // Teams / SharePoint URL gate — only applied to the transcript flows.
  function isTranscriptUrl(url) {
    if (!url) return false;
    try {
      const u = new URL(url);
      if (
        u.hostname === "teams.microsoft.com" ||
        u.hostname.endsWith(".teams.microsoft.com")
      ) {
        return true;
      }
      if (u.hostname.endsWith(".sharepoint.com")) {
        return true;
      }
      return false;
    } catch (_) {
      return false;
    }
  }

  // Reject tabs we cannot script or capture (chrome:// pages, extension
  // pages, the Chrome Web Store, view-source:, etc). Used by the PNG flow.
  function isCapturableUrl(url) {
    if (!url) return false;
    try {
      const u = new URL(url);
      if (u.protocol !== "http:" && u.protocol !== "https:" && u.protocol !== "file:") {
        return false;
      }
      if (u.hostname === "chrome.google.com") return false;
      return true;
    } catch (_) {
      return false;
    }
  }

  function closePort() {
    if (activePort) {
      try {
        activePort.disconnect();
      } catch (_) {
        /* ignore */
      }
      activePort = null;
    }
    activeCancel = null;
    setBusy(false);
  }

  // ---------------------------------------------------------------------------
  // Transcript flows (Quick Grab + Grab as MD/TXT/PDF)
  // ---------------------------------------------------------------------------

  async function startTranscriptExport(format) {
    setStatus("Starting...", "info");
    setDetail("");
    setBusy(true, false);

    try {
      const [tab] = await chrome.tabs.query({
        active: true,
        currentWindow: true,
      });

      if (!tab || !tab.id) {
        setStatus("No active tab found.", "error");
        setBusy(false);
        return;
      }

      if (!isTranscriptUrl(tab.url)) {
        setStatus(
          "Open a Teams or SharePoint Stream recording page, then try again.",
          "error"
        );
        setBusy(false);
        return;
      }

      // Inject the content script into every frame so we can scrape either
      // the top SharePoint Stream frame or the Teams Recap xplat iframe.
      await chrome.scripting.executeScript({
        target: { tabId: tab.id, allFrames: true },
        files: ["content.js"],
      });

      const targetFrameId = await pickTranscriptFrame(tab.id);
      const parentMetadata =
        targetFrameId !== 0 ? await probeParentMetadata(tab.id) : null;

      const port = chrome.tabs.connect(tab.id, {
        name: "meetmark",
        frameId: targetFrameId,
      });
      activePort = port;
      activeCancel = () => {
        try {
          port.postMessage({ type: "cancel" });
        } catch (_) {}
      };

      port.onDisconnect.addListener(() => {
        if (activePort === port) activePort = null;
      });

      port.onMessage.addListener((msg) => {
        if (!msg || !msg.type) return;
        switch (msg.type) {
          case "progress":
            setStatus(msg.message || "Working...", "info");
            setDetail(msg.detail || "");
            break;
          case "warning":
            setStatus(msg.message || "Warning", "warn");
            break;
          case "done":
            handleTranscriptDone(msg);
            break;
          case "error":
            setStatus(msg.message || "Export failed.", "error");
            setDetail("");
            closePort();
            break;
          default:
            break;
        }
      });

      port.postMessage({ type: "start", format, parentMetadata });
      setStatus("Reading transcript...", "info");
    } catch (err) {
      setStatus(
        "Error: " + (err && err.message ? err.message : String(err)),
        "error"
      );
      closePort();
    }
  }

  async function handleTranscriptDone(msg) {
    setStatus("Converting and downloading...", "info");
    setDetail("");
    try {
      const downloadResult = await chrome.runtime.sendMessage({
        type: "MEETMARK_DOWNLOAD",
        filename: msg.filename,
        content: msg.content,
        mime: msg.mime,
        format: msg.format,
      });

      if (downloadResult && downloadResult.ok) {
        const what =
          msg.format === "pdf"
            ? "Opened print-to-PDF view: " + msg.filename
            : "Done. Downloaded as " + msg.filename;
        if (msg.warning) {
          setStatus("Exported with warning: " + msg.warning, "warn");
        } else {
          setStatus(what, "success");
        }
        setDetail(
          (msg.turnCount ? msg.turnCount + " turns exported" : "") +
            (msg.speakerCount
              ? (msg.turnCount ? ", " : "") + msg.speakerCount + " speakers"
              : "")
        );
      } else {
        setStatus(
          "Download failed: " +
            ((downloadResult && downloadResult.error) || "unknown error"),
          "error"
        );
      }
    } catch (err) {
      setStatus(
        "Download error: " + (err && err.message ? err.message : String(err)),
        "error"
      );
    } finally {
      closePort();
    }
  }

  // ---------------------------------------------------------------------------
  // PNG full-page screenshot flow
  // ---------------------------------------------------------------------------
  //
  // The popup owns the chrome.tabs.captureVisibleTab calls and the canvas
  // stitching. A content script (screenshot.js) injected into the top frame
  // owns the scroll loop — it tells us which (x, y) to capture at each step
  // and waits for our acknowledgement before scrolling to the next.
  //
  // Chrome throttles captureVisibleTab to ~2/second; screenshot.js already
  // waits CAPTURE_DELAY between scrolls, but we also rate-limit here by
  // awaiting each capture before replying.

  async function startScreenshot() {
    setStatus("Starting screenshot...", "info");
    setDetail("");
    setBusy(true, true);
    setProgress(0);

    let tab;
    try {
      [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    } catch (err) {
      setStatus(
        "Error: " + (err && err.message ? err.message : String(err)),
        "error"
      );
      setBusy(false);
      return;
    }

    if (!tab || !tab.id) {
      setStatus("No active tab found.", "error");
      setBusy(false);
      return;
    }

    if (!isCapturableUrl(tab.url)) {
      setStatus(
        "This page can't be captured (chrome:// pages, the Chrome Web Store, and extension pages are blocked by the browser).",
        "error"
      );
      setBusy(false);
      return;
    }

    try {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id, frameIds: [0] },
        files: ["screenshot.js"],
      });
    } catch (err) {
      setStatus(
        "Could not attach to page: " +
          (err && err.message ? err.message : String(err)),
        "error"
      );
      setBusy(false);
      return;
    }

    // Canvas(es) we stitch captures into. Lazily created once we know the
    // page's full dimensions and the devicePixelRatio scale from the first
    // captured image.
    let screenshots = [];
    let lastCaptureError = null;

    const port = chrome.tabs.connect(tab.id, {
      name: "meetmark-shot",
      frameId: 0,
    });
    activePort = port;
    activeCancel = () => {
      try {
        port.postMessage({ type: "cancel" });
      } catch (_) {}
    };

    port.onDisconnect.addListener(() => {
      if (activePort === port) activePort = null;
    });

    port.onMessage.addListener(async (msg) => {
      if (!msg || !msg.type) return;
      try {
        if (msg.type === "capture") {
          await handleCaptureAt(msg, screenshots, tab.id);
          setProgress(msg.complete || 0);
          try {
            port.postMessage({ type: "captured" });
          } catch (_) {}
        } else if (msg.type === "done") {
          await finishScreenshot(screenshots, tab);
        } else if (msg.type === "error") {
          setStatus(msg.message || "Screenshot failed.", "error");
          closePort();
        } else if (msg.type === "progress") {
          setStatus(msg.message || "Working...", "info");
        }
      } catch (err) {
        lastCaptureError = err;
        setStatus(
          "Screenshot error: " +
            (err && err.message ? err.message : String(err)),
          "error"
        );
        try {
          port.postMessage({ type: "cancel" });
        } catch (_) {}
        closePort();
      }
    });

    setStatus("Measuring page...", "info");
    port.postMessage({ type: "start" });

    // If screenshot.js never responds (e.g. a weird CSP blocks injection),
    // the port will idle indefinitely. Give up after 30 s with no messages.
    // (Each capture resets the timer via the message listener above — we
    // don't wire that up for simplicity; the overall 30 s is a backstop.)
    setTimeout(() => {
      if (activePort === port && !lastCaptureError) {
        // Still running — don't cancel. This handler is only for the case
        // where screenshot.js never sent the first message. Best-effort
        // heuristic: if no captures happened yet, bail.
        if (screenshots.length === 0) {
          setStatus("Page did not respond. Try reloading and retrying.", "error");
          try {
            port.postMessage({ type: "cancel" });
          } catch (_) {}
          closePort();
        }
      }
    }, 30000);
  }

  // Draw one viewport capture into the appropriate stitched canvas(es).
  async function handleCaptureAt(info, screenshots, tabId) {
    // We call captureVisibleTab without a windowId — it captures the active
    // tab of the current window, which is the tab under the popup.
    const dataUrl = await chrome.tabs.captureVisibleTab(null, {
      format: "png",
    });
    if (!dataUrl) {
      throw new Error("captureVisibleTab returned no data. The browser may have throttled the capture.");
    }

    const image = await loadImage(dataUrl);

    // On high-DPI displays captureVisibleTab returns a larger image than the
    // logical window size. Scale the stitching math to match.
    let x = info.x;
    let y = info.y;
    let totalWidth = info.totalWidth;
    let totalHeight = info.totalHeight;
    if (info.windowWidth && info.windowWidth !== image.width) {
      const scale = image.width / info.windowWidth;
      x *= scale;
      y *= scale;
      totalWidth *= scale;
      totalHeight *= scale;
    }

    // Lazily init the canvas(es) once we know the real pixel dimensions.
    if (!screenshots.length) {
      initScreenshots(totalWidth, totalHeight).forEach((s) =>
        screenshots.push(s)
      );
    }

    // Draw into every canvas whose rect intersects this capture's rect.
    const imgRight = x + image.width;
    const imgBottom = y + image.height;
    for (const s of screenshots) {
      if (
        x < s.right &&
        imgRight > s.left &&
        y < s.bottom &&
        imgBottom > s.top
      ) {
        s.ctx.drawImage(image, x - s.left, y - s.top);
      }
    }
  }

  // Max canvas dimension Chrome will reliably encode to PNG. Matches
  // mrcoles' full-page-screen-capture constants.
  const MAX_PRIMARY = 15000 * 2;
  const MAX_SECONDARY = 4000 * 2;
  const MAX_AREA = MAX_PRIMARY * MAX_SECONDARY;

  // If the stitched image is bigger than the browser can safely encode to
  // PNG, break it up into tiles. Returns an array of { canvas, ctx, left,
  // top, right, bottom }.
  function initScreenshots(totalWidth, totalHeight) {
    const tooBig =
      totalHeight > MAX_PRIMARY ||
      totalWidth > MAX_PRIMARY ||
      totalHeight * totalWidth > MAX_AREA;
    const biggerWidth = totalWidth > totalHeight;
    const maxWidth = !tooBig
      ? totalWidth
      : biggerWidth
      ? MAX_PRIMARY
      : MAX_SECONDARY;
    const maxHeight = !tooBig
      ? totalHeight
      : biggerWidth
      ? MAX_SECONDARY
      : MAX_PRIMARY;
    const numCols = Math.max(1, Math.ceil(totalWidth / maxWidth));
    const numRows = Math.max(1, Math.ceil(totalHeight / maxHeight));

    const result = [];
    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        const canvas = document.createElement("canvas");
        canvas.width =
          col === numCols - 1 ? totalWidth - col * maxWidth : maxWidth;
        canvas.height =
          row === numRows - 1 ? totalHeight - row * maxHeight : maxHeight;
        const left = col * maxWidth;
        const top = row * maxHeight;
        result.push({
          canvas,
          ctx: canvas.getContext("2d"),
          left,
          top,
          right: left + canvas.width,
          bottom: top + canvas.height,
        });
      }
    }
    return result;
  }

  function loadImage(dataUrl) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () =>
        reject(new Error("Could not decode captured viewport image."));
      img.src = dataUrl;
    });
  }

  async function finishScreenshot(screenshots, tab) {
    if (!screenshots.length) {
      setStatus("Nothing to save — no viewport captures collected.", "error");
      closePort();
      return;
    }

    setStatus("Encoding PNG...", "info");
    setProgress(1);

    const baseName = buildScreenshotFilename(tab.url);
    let downloaded = 0;
    const total = screenshots.length;

    for (let i = 0; i < screenshots.length; i++) {
      const suffix = total > 1 ? "-" + (i + 1) : "";
      const filename = baseName + suffix + ".png";
      const dataUrl = screenshots[i].canvas.toDataURL("image/png");

      const result = await chrome.runtime.sendMessage({
        type: "MEETMARK_DOWNLOAD",
        filename,
        dataUrl,
        mime: "image/png",
        format: "png",
      });

      if (result && result.ok) {
        downloaded++;
      } else {
        setStatus(
          "Download failed: " +
            ((result && result.error) || "unknown error"),
          "error"
        );
        closePort();
        return;
      }
    }

    if (total > 1) {
      setStatus(
        "Done. Page was too tall for a single image — saved " +
          total +
          " PNG tiles.",
        "success"
      );
    } else {
      setStatus("Done. Saved full-page PNG (" + downloaded + " file).", "success");
    }
    setDetail("");
    closePort();
  }

  function buildScreenshotFilename(url) {
    let base = "screenshot";
    try {
      const u = new URL(url);
      base =
        (u.hostname + u.pathname)
          .replace(/[\\/:*?"<>|]/g, "_")
          .replace(/_+$/g, "")
          .slice(0, 80) || "screenshot";
    } catch (_) {}
    const d = new Date();
    const stamp =
      d.getFullYear() +
      "." +
      String(d.getMonth() + 1).padStart(2, "0") +
      "." +
      String(d.getDate()).padStart(2, "0") +
      "-" +
      String(d.getHours()).padStart(2, "0") +
      String(d.getMinutes()).padStart(2, "0") +
      String(d.getSeconds()).padStart(2, "0");
    return stamp + " - " + base;
  }

  // ---------------------------------------------------------------------------
  // Button wiring
  // ---------------------------------------------------------------------------

  quickBtn.addEventListener("click", () => startTranscriptExport("md"));
  grabBtn.addEventListener("click", () =>
    startTranscriptExport(formatSelect.value || "md")
  );
  pngBtn.addEventListener("click", () => startScreenshot());

  cancelBtn.addEventListener("click", () => {
    if (activeCancel) {
      try {
        activeCancel();
      } catch (_) {}
    }
    closePort();
    window.close();
  });

  // ---------------------------------------------------------------------------
  // Helpers for frame selection (unchanged from the prior single-purpose flow)
  // ---------------------------------------------------------------------------

  async function pickTranscriptFrame(tid) {
    let probes;
    try {
      probes = await chrome.scripting.executeScript({
        target: { tabId: tid, allFrames: true },
        func: () => {
          try {
            return {
              url: location.href,
              hostname: location.hostname,
              hasStreamCells:
                document.querySelector('[data-automationid="ListCell"]') !==
                null,
              hasTranscriptHooks:
                document.querySelector(
                  '[data-tid="Transcript"],[aria-label*="ranscript" i],[data-automationid*="transcript" i]'
                ) !== null,
            };
          } catch (_) {
            return {
              url: "",
              hostname: "",
              hasStreamCells: false,
              hasTranscriptHooks: false,
            };
          }
        },
      });
    } catch (_) {
      return 0;
    }

    if (!probes || !probes.length) return 0;

    const cellFrame = probes.find((p) => p.result && p.result.hasStreamCells);
    if (cellFrame) return cellFrame.frameId;

    const hookFrame = probes.find(
      (p) => p.result && p.result.hasTranscriptHooks
    );
    if (hookFrame) return hookFrame.frameId;

    const top = probes.find((p) => p.frameId === 0);
    const topHost = (top && top.result && top.result.hostname) || "";
    if (/(^|\.)teams\.microsoft\.com$/i.test(topHost)) {
      const spFrame = probes.find(
        (p) =>
          p.frameId !== 0 &&
          p.result &&
          /\.sharepoint\.com$/i.test(p.result.hostname || "")
      );
      if (spFrame) return spFrame.frameId;
    }

    return 0;
  }

  async function probeParentMetadata(tid) {
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId: tid, frameIds: [0] },
        func: () => {
          function stripTitleNoise(t) {
            if (!t) return "";
            let out = t;
            out = out.replace(
              /\s*[|\-]\s*(Microsoft Teams|Microsoft Stream|SharePoint|OneDrive).*$/i,
              ""
            );
            out = out.replace(/\.[a-z0-9]{2,5}$/i, "");
            out = out.replace(/[-_\s]+Meeting Recording\s*$/i, "");
            out = out.replace(/[-_\s]+\d{8}[_-]?\d{0,6}.*$/i, "");
            return out.replace(/\s+/g, " ").trim();
          }

          let title = "";
          const titleSelectors = [
            '[data-tid="entity-header"] .fui-StyledText',
            '[data-tid="app-layout-area--header"] .fui-StyledText',
            '[data-tid="entity-header"] span[dir="auto"]',
            '[data-automationid="videoTitle"]',
            '[data-automationid="video-title"]',
            '[data-automationid="pageTitle"]',
            "main h1",
            "h1",
          ];
          for (const sel of titleSelectors) {
            const el = document.querySelector(sel);
            if (!el) continue;
            const t = (el.innerText || el.textContent || "")
              .replace(/\s+/g, " ")
              .trim();
            if (t && t.length >= 2 && t.length < 200) {
              title = stripTitleNoise(t);
              break;
            }
          }
          if (!title) {
            title = stripTitleNoise((document.title || "").trim());
          }

          let dateIso = "";
          const monthPattern =
            /(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})/i;
          const scopes = [document.querySelector("main"), document.body].filter(
            Boolean
          );
          for (const scope of scopes) {
            const text = (scope.innerText || "").slice(0, 20000);
            const m = text.match(monthPattern);
            if (m) {
              const parsed = new Date(m[0]);
              if (!isNaN(parsed.getTime())) {
                dateIso = parsed.toISOString();
                break;
              }
            }
          }

          return {
            title,
            dateIso,
            url: location.href,
          };
        },
      });

      if (!results || !results.length || !results[0].result) return null;
      return results[0].result;
    } catch (_) {
      return null;
    }
  }
});

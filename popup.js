// MeetMark popup controller.
//
// The popup exposes two export paths:
//
//   1. Grab transcript as (MD / TXT / PDF) — runs the Teams / SharePoint
//      Stream transcript pipeline and re-renders the result in the chosen
//      format. Markdown is the default option, and the "Grab" button is
//      effectively still one click for that case. PDF opens a printable
//      HTML tab so the user can Save as PDF from the browser print dialog.
//
//   2. Capture full page as PNG — full-page screenshot of the current tab.
//      We use the Chrome DevTools Protocol via chrome.debugger:
//      Page.getLayoutMetrics tells us the real document size, then
//      Page.captureScreenshot({ captureBeyondViewport: true, ... }) returns
//      one PNG spanning the whole document. This is how GoFullPage works
//      and it correctly captures pages where the content lives inside an
//      inner scrollable container (dashboards, Fluent UI panels, etc.) —
//      situations where plain window.scrollTo() does nothing.
//      Works on any http(s) tab (relies on activeTab + debugger).

document.addEventListener("DOMContentLoaded", () => {
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
  // We use the Chrome DevTools Protocol via chrome.debugger:
  //
  //   Page.getLayoutMetrics  → the real document size, including content
  //                            that lives inside inner scrollable elements
  //                            (dashboards, Fluent UI panels, etc.) where
  //                            window.scrollTo() does not work.
  //   Page.captureScreenshot({ captureBeyondViewport: true, clip })
  //                          → one PNG spanning the full document.
  //
  // This replaces the earlier scroll-and-stitch approach, which only worked
  // when window was the scrollable element and produced a single-viewport
  // PNG on pages with inner scroll containers.

  let debuggerAttachedTarget = null;

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

    if (!chrome.debugger || !chrome.debugger.attach) {
      setStatus(
        "Full-page PNG capture requires Chrome's debugger API, which is not available in this browser.",
        "error"
      );
      setBusy(false);
      return;
    }

    const target = { tabId: tab.id };
    activeCancel = async () => {
      await safeDetach(target);
    };

    try {
      setStatus("Attaching to page...", "info");
      await debuggerAttach(target, "1.3");
      debuggerAttachedTarget = target;

      setStatus("Measuring page...", "info");
      // Enable the Page domain so getLayoutMetrics returns stable values.
      await debuggerSend(target, "Page.enable", {});
      const metrics = await debuggerSend(
        target,
        "Page.getLayoutMetrics",
        {}
      );

      // cssContentSize is the post-CSS-pixel document size (what we want);
      // Chrome >=89 exposes it, older versions only expose contentSize in
      // device pixels. Fall back as needed.
      const size = metrics.cssContentSize || metrics.contentSize;
      if (!size || !size.width || !size.height) {
        throw new Error("Browser did not report a document size.");
      }
      const width = Math.max(1, Math.ceil(size.width));
      const height = Math.max(1, Math.ceil(size.height));

      setProgress(0.25);
      setStatus(
        "Capturing full page (" + width + " × " + height + ")...",
        "info"
      );

      // captureBeyondViewport: true tells Chrome to render the entire clip,
      // including offscreen/inner-scrollable content, into one image.
      const shot = await debuggerSend(target, "Page.captureScreenshot", {
        format: "png",
        captureBeyondViewport: true,
        fromSurface: true,
        clip: { x: 0, y: 0, width, height, scale: 1 },
      });

      if (!shot || !shot.data) {
        throw new Error("Capture returned no image data.");
      }

      setProgress(0.9);
      setStatus("Downloading...", "info");

      const filename = buildScreenshotFilename(tab.url) + ".png";
      const dataUrl = "data:image/png;base64," + shot.data;
      const dl = await chrome.runtime.sendMessage({
        type: "MEETMARK_DOWNLOAD",
        filename,
        dataUrl,
        mime: "image/png",
        format: "png",
      });

      if (dl && dl.ok) {
        setProgress(1);
        setStatus(
          "Done. Saved full-page PNG: " + filename,
          "success"
        );
        setDetail(width + " × " + height + " pixels");
      } else {
        setStatus(
          "Download failed: " + ((dl && dl.error) || "unknown error"),
          "error"
        );
      }
    } catch (err) {
      const msg = (err && err.message) || String(err);
      if (/cannot access|debugger/i.test(msg) && /another/i.test(msg)) {
        setStatus(
          "Another debugger is already attached to this tab (DevTools open?). Close DevTools and try again.",
          "error"
        );
      } else {
        setStatus("Screenshot error: " + msg, "error");
      }
    } finally {
      await safeDetach(target);
      activeCancel = null;
      setBusy(false);
    }
  }

  // Promisified wrappers — chrome.debugger still uses callbacks in MV3.
  function debuggerAttach(target, version) {
    return new Promise((resolve, reject) => {
      try {
        chrome.debugger.attach(target, version, () => {
          const err = chrome.runtime.lastError;
          if (err) return reject(new Error(err.message));
          resolve();
        });
      } catch (e) {
        reject(e);
      }
    });
  }

  function debuggerSend(target, method, params) {
    return new Promise((resolve, reject) => {
      try {
        chrome.debugger.sendCommand(target, method, params || {}, (result) => {
          const err = chrome.runtime.lastError;
          if (err) return reject(new Error(err.message));
          resolve(result);
        });
      } catch (e) {
        reject(e);
      }
    });
  }

  async function safeDetach(target) {
    if (!debuggerAttachedTarget) return;
    debuggerAttachedTarget = null;
    try {
      await new Promise((resolve) => {
        chrome.debugger.detach(target, () => {
          // Swallow lastError — we don't care if the target is already gone.
          void chrome.runtime.lastError;
          resolve();
        });
      });
    } catch (_) {
      /* ignore */
    }
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

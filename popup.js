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
//      Faithful port of mrcoles / GoFullPage's approach: inject a content
//      script that scrolls the page through a grid of viewport positions
//      and, for each, calls chrome.tabs.captureVisibleTab. The popup
//      stitches the resulting images into one (or more, for very long
//      pages) offscreen canvases and saves the PNG(s) via chrome.downloads.
//      Works on any http(s) tab; only activeTab is required.

document.addEventListener("DOMContentLoaded", () => {
  const grabBtn = document.getElementById("grabBtn");
  const pngBtn = document.getElementById("pngBtn");
  const formatSelect = document.getElementById("formatSelect");
  const cancelBtn = document.getElementById("cancelBtn");
  const cancelRow = document.getElementById("cancelRow");
  const status = document.getElementById("status");
  const detail = document.getElementById("detail");
  const progressBar = document.getElementById("progressBar");
  const progressFill = progressBar.querySelector(".fill");
  const dlAutoRadio = document.getElementById("dlAuto");
  const dlAskRadio = document.getElementById("dlAsk");
  const folderRow = document.getElementById("folderRow");
  const folderNameEl = document.getElementById("folderName");
  const chooseFolderBtn = document.getElementById("chooseFolderBtn");

  // Download preferences: "auto" (default) or "ask".
  let dlMode = "auto";
  // Persisted FileSystemDirectoryHandle from showDirectoryPicker(); null = use
  // the browser's default Downloads directory.
  let savedDirHandle = null;

  // ---------------------------------------------------------------------------
  // IndexedDB helpers for persisting the FileSystemDirectoryHandle
  // ---------------------------------------------------------------------------

  function openSettingsDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open("meetmark-settings", 1);
      req.onupgradeneeded = () => req.result.createObjectStore("kv");
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function idbGet(key) {
    const db = await openSettingsDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction("kv", "readonly");
      const req = tx.objectStore("kv").get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function idbSet(key, value) {
    const db = await openSettingsDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction("kv", "readwrite");
      tx.objectStore("kv").put(value, key);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  }

  // ---------------------------------------------------------------------------
  // Save-to-folder via File System Access API
  // ---------------------------------------------------------------------------

  // Write content directly to the user's chosen directory without going through
  // chrome.downloads. content can be a Blob or a string (mime required for string).
  async function saveToDir(handle, filename, content, mime) {
    const blob =
      content instanceof Blob
        ? content
        : new Blob([content], { type: mime || "application/octet-stream" });

    let perm = await handle.queryPermission({ mode: "readwrite" });
    if (perm === "prompt") {
      try {
        perm = await handle.requestPermission({ mode: "readwrite" });
      } catch (_) {
        perm = "denied";
      }
    }
    if (perm !== "granted") {
      throw new Error(
        'Permission to the chosen folder was denied. Click "Choose folder…" to reselect it.'
      );
    }

    const fileHandle = await handle.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
  }

  function applyDlMode(mode) {
    dlMode = mode;
    dlAutoRadio.checked = mode === "auto";
    dlAskRadio.checked = mode === "ask";
    folderRow.classList.toggle("hidden", mode !== "auto");
  }

  // Load persisted settings.
  chrome.storage.local.get(["dlMode"], (prefs) => {
    applyDlMode(prefs.dlMode === "ask" ? "ask" : "auto");
  });

  idbGet("dirHandle").then((handle) => {
    if (handle && handle.kind === "directory") {
      savedDirHandle = handle;
      folderNameEl.textContent = handle.name;
    }
  }).catch(() => {});

  dlAutoRadio.addEventListener("change", () => {
    if (dlAutoRadio.checked) {
      applyDlMode("auto");
      chrome.storage.local.set({ dlMode: "auto" });
    }
  });
  dlAskRadio.addEventListener("change", () => {
    if (dlAskRadio.checked) {
      applyDlMode("ask");
      chrome.storage.local.set({ dlMode: "ask" });
    }
  });

  chooseFolderBtn.addEventListener("click", async () => {
    try {
      const handle = await window.showDirectoryPicker({
        id: "meetmark-save",
        mode: "readwrite",
      });
      savedDirHandle = handle;
      folderNameEl.textContent = handle.name;
      await idbSet("dirHandle", handle);
    } catch (err) {
      if (err && err.name !== "AbortError") {
        setStatus("Could not access folder: " + err.message, "warn");
      }
    }
  });

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
    cancelRow.classList.toggle("visible", busy);
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

  // Teams/SharePoint host access is declared as optional so the extension
  // carries no persistent "Site access" warning. We request it on first use.
  async function ensureTranscriptPermissions() {
    const origins = [
      "https://teams.microsoft.com/*",
      "https://*.teams.microsoft.com/*",
      "https://*.sharepoint.com/*",
    ];
    try {
      const has = await new Promise((resolve) =>
        chrome.permissions.contains({ origins }, resolve)
      );
      if (has) return true;
      const granted = await new Promise((resolve) =>
        chrome.permissions.request({ origins }, resolve)
      );
      return !!granted;
    } catch (_) {
      return false;
    }
  }

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

      const permitted = await ensureTranscriptPermissions();
      if (!permitted) {
        setStatus(
          "Permission required to access Teams and SharePoint pages.",
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
      let ok = false;
      let errMsg = "";

      if (msg.format !== "pdf" && dlMode === "auto" && savedDirHandle) {
        // Write directly to the user's chosen folder.
        try {
          await saveToDir(savedDirHandle, msg.filename, msg.content, msg.mime);
          ok = true;
        } catch (dirErr) {
          errMsg = dirErr && dirErr.message ? dirErr.message : String(dirErr);
        }
      }

      if (!ok && errMsg === "") {
        // Fall back to chrome.downloads (also used for PDF and "ask" mode).
        const downloadResult = await chrome.runtime.sendMessage({
          type: "MEETMARK_DOWNLOAD",
          filename: msg.filename,
          content: msg.content,
          mime: msg.mime,
          format: msg.format,
          saveAs: dlMode === "ask",
        });
        ok = !!(downloadResult && downloadResult.ok);
        errMsg = (downloadResult && downloadResult.error) || "unknown error";
      }

      if (ok) {
        const what =
          msg.format === "pdf"
            ? "Opened print-to-PDF view: " + msg.filename
            : "Done. Saved as " + msg.filename;
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
        setStatus("Download failed: " + errMsg, "error");
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
  // Port of mrcoles / GoFullPage, extended to handle dashboards where the
  // real scrolling is done by an inner <div> and not the window. The popup
  // drives the capture:
  //
  //   1. Inject screenshot.js into the top frame.
  //   2. Open a long-lived port and tell it to start.
  //   3. The content script picks the scroll target (window or the biggest
  //      scrollable descendant), then scrolls through a grid of viewport
  //      positions (bottom-up, left-to-right). For each it posts a `capture`
  //      message containing the current scroll offset, the full scrollable
  //      size, and — when the target is an inner element — the container's
  //      rect on the captured tab image so we can crop to it.
  //   4. For each `capture` we call chrome.tabs.captureVisibleTab, load the
  //      PNG into an Image, crop to the reported rect, and draw it into one
  //      (or more, for very tall pages) offscreen canvases at the correct
  //      pixel offset, then ack the content script with `captured`.
  //   5. On `done`, encode each canvas to a PNG data URL and hand it to the
  //      service worker to save via chrome.downloads.
  //
  // Canvas tiling matches GoFullPage: anything beyond 30000 device px in
  // either axis splits into multiple output PNGs so we never hit the
  // browser's canvas size limits.

  // GoFullPage's canvas-size guard. If the full document exceeds this on
  // either axis (in device pixels) we split the output across several PNGs.
  const MAX_PRIMARY_DIMENSION = 15000 * 2;

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
        "Couldn't inject capture script: " +
          (err && err.message ? err.message : String(err)),
        "error"
      );
      setBusy(false);
      return;
    }

    const windowId = tab.windowId;
    const port = chrome.tabs.connect(tab.id, {
      name: "meetmark-shot",
      frameId: 0,
    });
    activePort = port;

    // Stitching state — populated on the first capture message.
    let screenshots = null; // array of { canvas, ctx, left, top, right, bottom }
    let pageWidthPx = 0;
    let pageHeightPx = 0;
    let cancelled = false;
    let captureInFlight = Promise.resolve();

    activeCancel = () => {
      cancelled = true;
      try {
        port.postMessage({ type: "cancel" });
      } catch (_) {}
    };

    port.onDisconnect.addListener(() => {
      if (activePort === port) activePort = null;
      if (!cancelled && !screenshots) {
        // Port closed before any capture arrived — script didn't load.
        setStatus(
          "Couldn't talk to the page. Reload the tab and try again.",
          "error"
        );
        setBusy(false);
      }
    });

    port.onMessage.addListener((msg) => {
      if (!msg || !msg.type) return;
      switch (msg.type) {
        case "progress":
          setStatus(msg.message || "Capturing...", "info");
          break;
        case "capture":
          // Serialize: captureVisibleTab is rate-limited and we also need
          // each drawImage to land before the next scroll.
          captureInFlight = captureInFlight.then(() =>
            handleCapture(msg).catch((err) => {
              if (cancelled) return;
              cancelled = true;
              setStatus(
                "Capture error: " +
                  (err && err.message ? err.message : String(err)),
                "error"
              );
              try {
                port.postMessage({ type: "cancel" });
              } catch (_) {}
              closePort();
            })
          );
          break;
        case "done":
          captureInFlight.then(() => handleDone());
          break;
        case "error":
          setStatus(msg.message || "Capture failed.", "error");
          cancelled = true;
          closePort();
          break;
        default:
          break;
      }
    });

    async function handleCapture(data) {
      if (cancelled) return;

      const dataUrl = await captureVisibleTabWithRetry(windowId);
      if (cancelled) return;

      const img = await loadImage(dataUrl);
      if (cancelled) return;

      if (!screenshots) {
        pageWidthPx = Math.max(
          1,
          Math.round(data.totalWidth * data.devicePixelRatio)
        );
        pageHeightPx = Math.max(
          1,
          Math.round(data.totalHeight * data.devicePixelRatio)
        );
        screenshots = initScreenshots(pageWidthPx, pageHeightPx);
      }

      // captureVisibleTab returns at device-pixel resolution. When the
      // content script drove window.scrollTo the whole image is the content
      // and data.src is null. When it drove an inner scroll container,
      // data.src describes the container's rect on the tab image (CSS px)
      // and we need to crop before stitching.
      const dpr = data.devicePixelRatio || 1;

      let srcX = 0;
      let srcY = 0;
      let srcW = img.width;
      let srcH = img.height;
      if (data.src) {
        srcX = Math.max(0, Math.round(data.src.left * dpr));
        srcY = Math.max(0, Math.round(data.src.top * dpr));
        srcW = Math.max(0, Math.round(data.src.width * dpr));
        srcH = Math.max(0, Math.round(data.src.height * dpr));
        // Clamp to image bounds.
        if (srcX >= img.width || srcY >= img.height) {
          // Container is fully off-screen; nothing to draw.
          if (typeof data.complete === "number") setProgress(data.complete);
          try {
            port.postMessage({ type: "captured" });
          } catch (_) {}
          return;
        }
        srcW = Math.min(srcW, img.width - srcX);
        srcH = Math.min(srcH, img.height - srcY);
      }

      if (srcW <= 0 || srcH <= 0) {
        if (typeof data.complete === "number") setProgress(data.complete);
        try {
          port.postMessage({ type: "captured" });
        } catch (_) {}
        return;
      }

      // Destination in the output canvas (device px). data.x / y are the
      // scroll offset inside the container (or the window).
      const dstX = Math.round(data.x * dpr);
      const dstY = Math.round(data.y * dpr);
      const dstRight = dstX + srcW;
      const dstBottom = dstY + srcH;

      for (const s of screenshots) {
        if (dstRight <= s.left || s.right <= dstX) continue;
        if (dstBottom <= s.top || s.bottom <= dstY) continue;
        s.ctx.drawImage(
          img,
          srcX,
          srcY,
          srcW,
          srcH,
          dstX - s.left,
          dstY - s.top,
          srcW,
          srcH
        );
      }

      if (typeof data.complete === "number") {
        setProgress(data.complete);
      }

      try {
        port.postMessage({ type: "captured" });
      } catch (_) {}
    }

    async function handleDone() {
      if (cancelled) {
        closePort();
        return;
      }
      if (!screenshots || !screenshots.length) {
        setStatus("Capture produced no image data.", "error");
        closePort();
        return;
      }

      try {
        setStatus("Encoding PNG...", "info");
        const base = buildScreenshotFilename(tab.url);
        const multi = screenshots.length > 1;
        let savedCount = 0;
        let lastError = "";

        for (let i = 0; i < screenshots.length; i++) {
          const filename = multi
            ? base + "-part-" + (i + 1) + "-of-" + screenshots.length + ".png"
            : base + ".png";
          const blob = await canvasToBlob(screenshots[i].canvas);

          if (dlMode === "auto" && savedDirHandle) {
            try {
              await saveToDir(savedDirHandle, filename, blob, "image/png");
              savedCount++;
            } catch (dirErr) {
              lastError = dirErr && dirErr.message ? dirErr.message : String(dirErr);
            }
          } else {
            const dataUrl = await blobToDataUrl(blob);
            const dl = await chrome.runtime.sendMessage({
              type: "MEETMARK_DOWNLOAD",
              filename,
              dataUrl,
              mime: "image/png",
              format: "png",
              saveAs: dlMode === "ask",
            });
            if (dl && dl.ok) {
              savedCount++;
            } else {
              lastError = (dl && dl.error) || "unknown error";
            }
          }
        }

        if (savedCount === screenshots.length) {
          setProgress(1);
          setStatus(
            multi
              ? "Done. Saved " + savedCount + " PNG tiles."
              : "Done. Saved full-page PNG.",
            "success"
          );
          setDetail(pageWidthPx + " × " + pageHeightPx + " pixels");
        } else if (savedCount > 0) {
          setStatus(
            "Saved " +
              savedCount +
              " of " +
              screenshots.length +
              " tiles. Last error: " +
              lastError,
            "warn"
          );
        } else {
          setStatus("Download failed: " + lastError, "error");
        }
      } catch (err) {
        setStatus(
          "Encoding error: " +
            (err && err.message ? err.message : String(err)),
          "error"
        );
      } finally {
        closePort();
      }
    }

    try {
      port.postMessage({ type: "start" });
      setStatus("Scrolling and capturing...", "info");
    } catch (err) {
      setStatus(
        "Couldn't start capture: " +
          (err && err.message ? err.message : String(err)),
        "error"
      );
      closePort();
    }
  }

  // Build the per-tile canvases. GoFullPage splits at MAX_PRIMARY_DIMENSION
  // on each axis so we never exceed the browser's canvas size cap, and emits
  // one PNG per tile.
  function initScreenshots(pixelWidth, pixelHeight) {
    const xs = sliceAxis(pixelWidth);
    const ys = sliceAxis(pixelHeight);
    const result = [];
    for (let i = 0; i < ys.length - 1; i++) {
      for (let j = 0; j < xs.length - 1; j++) {
        const left = xs[j];
        const right = xs[j + 1];
        const top = ys[i];
        const bottom = ys[i + 1];
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(1, right - left);
        canvas.height = Math.max(1, bottom - top);
        result.push({
          canvas,
          ctx: canvas.getContext("2d"),
          left,
          top,
          right,
          bottom,
        });
      }
    }
    return result;
  }

  function sliceAxis(total) {
    const out = [0];
    let p = 0;
    while (p < total) {
      const next = Math.min(total, p + MAX_PRIMARY_DIMENSION);
      out.push(next);
      p = next;
    }
    if (out.length === 1) out.push(total); // zero-length page guard
    return out;
  }

  function loadImage(dataUrl) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () =>
        reject(new Error("Failed to decode captured image."));
      img.src = dataUrl;
    });
  }

  function canvasToBlob(canvas) {
    return new Promise((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (!blob) return reject(new Error("Canvas toBlob returned null."));
        resolve(blob);
      }, "image/png");
    });
  }

  function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(new Error("Blob read failed."));
      r.readAsDataURL(blob);
    });
  }

  // chrome.tabs.captureVisibleTab is throttled to roughly 2 calls/second
  // per window. The content script already awaits each ack before scrolling,
  // but we still add a small retry for the cases where Chrome reports
  // MAX_CAPTURE_VISIBLE_TAB_CALLS_PER_SECOND.
  function captureVisibleTabWithRetry(windowId) {
    let attempt = 0;
    function once() {
      return new Promise((resolve, reject) => {
        try {
          chrome.tabs.captureVisibleTab(
            windowId,
            { format: "png" },
            (url) => {
              const err = chrome.runtime.lastError;
              if (err) return reject(new Error(err.message));
              if (!url) return reject(new Error("Empty capture."));
              resolve(url);
            }
          );
        } catch (e) {
          reject(e);
        }
      });
    }
    return (function loop() {
      return once().catch((err) => {
        const m = (err && err.message) || String(err);
        if (attempt < 5 && /MAX_CAPTURE|per second|too many/i.test(m)) {
          attempt++;
          return new Promise((r) => setTimeout(r, 600)).then(loop);
        }
        throw err;
      });
    })();
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

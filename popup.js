// MeetMark popup controller.
// The export starts automatically when the popup opens. The Cancel button
// aborts immediately and closes the popup. The Export button retries after
// a completed or failed run.

document.addEventListener("DOMContentLoaded", () => {
  const exportBtn = document.getElementById("exportBtn");
  const cancelBtn = document.getElementById("cancelBtn");
  const status = document.getElementById("status");
  const detail = document.getElementById("detail");
  const progressBar = document.getElementById("progressBar");

  let port = null;
  let tabId = null;
  let targetFrameId = 0;

  function setStatus(message, level) {
    status.textContent = message || "";
    status.className = "status " + (level || "info");
  }

  function setDetail(text) {
    detail.textContent = text || "";
  }

  function setBusy(busy) {
    exportBtn.disabled = busy;
    cancelBtn.disabled = !busy;
    if (busy) {
      progressBar.classList.add("active");
    } else {
      progressBar.classList.remove("active");
    }
  }

  function isSupportedUrl(url) {
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

  // Tear down the port and reset the UI to idle (export button re-enabled).
  function closePort() {
    if (port) {
      try {
        port.disconnect();
      } catch (_) {
        /* ignore */
      }
      port = null;
    }
    setBusy(false);
  }

  // Start or re-run the export. Called automatically on open and by the
  // Export button when retrying after a completed or failed run.
  async function startExport() {
    setStatus("Starting...", "info");
    setDetail("");
    setBusy(true);

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

      if (!isSupportedUrl(tab.url)) {
        setStatus(
          "Open a Teams or SharePoint Stream recording page, then click Export again.",
          "error"
        );
        setBusy(false);
        return;
      }

      tabId = tab.id;

      // Inject the content script into every frame of the tab. On a
      // SharePoint Stream page the transcript lives in the top frame, but on
      // the teams.microsoft.com Calendar / Recap view the transcript is
      // rendered inside a cross-origin SharePoint iframe (xplatIframe). We
      // need the content script running in whichever frame owns the DOM.
      await chrome.scripting.executeScript({
        target: { tabId: tab.id, allFrames: true },
        files: ["content.js"],
      });

      // Probe every frame to find which one actually hosts the transcript.
      // Prefer a frame that already has transcript list cells rendered, then
      // one that has transcript panel hooks, then fall back to a SharePoint
      // iframe under a teams.microsoft.com parent, and finally the top frame.
      targetFrameId = await pickTranscriptFrame(tab.id);

      // When the scrape is happening inside a child frame, the Teams page
      // chrome — meeting title, date, page URL — lives in the top frame and
      // isn't reachable from the scraping frame. Probe the top frame for
      // those values so the content script can use them when formatting the
      // filename and the Markdown header.
      const parentMetadata =
        targetFrameId !== 0 ? await probeParentMetadata(tab.id) : null;

      // Open a long-lived Port to the content script in the chosen frame so
      // we can stream progress and send cancel.
      port = chrome.tabs.connect(tab.id, {
        name: "meetmark",
        frameId: targetFrameId,
      });

      port.onDisconnect.addListener(() => {
        port = null;
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
            handleDone(msg);
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

      port.postMessage({ type: "start", parentMetadata });
      setStatus("Reading transcript...", "info");
    } catch (err) {
      setStatus(
        "Error: " + (err && err.message ? err.message : String(err)),
        "error"
      );
      closePort();
    }
  }

  // Export / retry button.
  exportBtn.addEventListener("click", () => startExport());

  // Cancel: signal the content script, disconnect immediately, and close the
  // popup so the user isn't left staring at a frozen UI.
  cancelBtn.addEventListener("click", () => {
    if (port) {
      try {
        port.postMessage({ type: "cancel" });
      } catch (_) {
        /* ignore */
      }
    }
    closePort();
    window.close();
  });

  // Handle a "done" message: push markdown to the service worker for
  // download, then update status.
  async function handleDone(msg) {
    setStatus("Converting and downloading...", "info");
    setDetail("");
    try {
      const downloadResult = await chrome.runtime.sendMessage({
        type: "MEETMARK_DOWNLOAD",
        filename: msg.filename,
        markdown: msg.markdown,
      });

      if (downloadResult && downloadResult.ok) {
        if (msg.warning) {
          setStatus("Exported with warning: " + msg.warning, "warn");
        } else {
          setStatus("Done. Downloaded as " + msg.filename, "success");
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

  // Probe every frame in the tab and decide which one the content script
  // should scrape. Returns a frameId (0 = top). The heuristic, in order:
  //   1. A frame whose DOM already contains Stream transcript ListCells.
  //   2. A frame that has Teams/Stream transcript hooks (aria-label,
  //      data-tid="Transcript", data-automationid*="transcript").
  //   3. If the top frame is on teams.microsoft.com, the first child frame
  //      on *.sharepoint.com — the Teams Recap xplat iframe loads from
  //      SharePoint and is where the transcript DOM actually lives.
  //   4. The top frame.
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

  // Probe the top frame for the meeting title, date, and URL. Used when the
  // scrape is happening inside a cross-origin child frame (the Teams Recap
  // xplat iframe) because that frame can't see the Teams page chrome where
  // these values live. Returns { title, dateIso, url } or null on failure.
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

  // Begin exporting as soon as the popup opens — no button click required.
  startExport();
});

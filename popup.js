// MeetMark popup controller.
// Talks to the content script over a long-lived Port so we can stream
// progress updates back to the popup UI and send a Cancel signal while
// work is in flight.

document.addEventListener("DOMContentLoaded", () => {
  const exportBtn = document.getElementById("exportBtn");
  const cancelBtn = document.getElementById("cancelBtn");
  const status = document.getElementById("status");
  const detail = document.getElementById("detail");
  const progressBar = document.getElementById("progressBar");

  let port = null;
  let tabId = null;

  // Render a message in the popup status area with a severity class.
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

  // Check whether a URL looks like a page we can scrape: either a Microsoft
  // Teams transcript page, or a SharePoint Stream recording page.
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

  // Cleanly tear down a port connection and reset the UI to idle.
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

  // Wire up the Export button.
  exportBtn.addEventListener("click", async () => {
    setStatus("");
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
          "This tab isn't a Teams or SharePoint Stream page. Open a Teams recording and try again.",
          "error"
        );
        setBusy(false);
        return;
      }

      tabId = tab.id;

      // Inject the content script on demand. Manifest V3 requires
      // scripting.executeScript rather than a declared content_script.
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ["content.js"],
      });

      // Open a long-lived Port to the content script so we can stream
      // progress and send cancel.
      port = chrome.tabs.connect(tab.id, { name: "meetmark" });

      port.onDisconnect.addListener(() => {
        // If the content script goes away unexpectedly, treat it as an
        // error only if we never heard a "done" / "error".
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

      port.postMessage({ type: "start" });
      setStatus("Starting export...", "info");
    } catch (err) {
      setStatus("Error: " + (err && err.message ? err.message : String(err)), "error");
      closePort();
    }
  });

  // Wire up the Cancel button.
  cancelBtn.addEventListener("click", () => {
    if (!port) return;
    setStatus("Cancelling...", "warn");
    try {
      port.postMessage({ type: "cancel" });
    } catch (_) {
      /* ignore */
    }
  });

  // Handle a "done" message from the content script: push the markdown to
  // the service worker for download, then update status.
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
      setStatus("Download error: " + (err && err.message ? err.message : String(err)), "error");
    } finally {
      closePort();
    }
  }
});

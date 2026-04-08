// MeetMark popup controller.
// Handles the single "Export transcript" button, injects the content script
// into the active Teams tab, and relays results/errors to the user.

// Wait for the popup DOM before wiring up handlers.
document.addEventListener("DOMContentLoaded", () => {
  const button = document.getElementById("exportBtn");
  const status = document.getElementById("status");

  // Render a message in the popup status area with a severity class.
  function setStatus(message, level) {
    status.textContent = message;
    status.className = "status " + (level || "info");
  }

  // Check whether a URL looks like a Microsoft Teams page we can scrape.
  function isTeamsUrl(url) {
    if (!url) return false;
    try {
      const u = new URL(url);
      return (
        u.hostname === "teams.microsoft.com" ||
        u.hostname.endsWith(".teams.microsoft.com")
      );
    } catch (_) {
      return false;
    }
  }

  // Main click handler: find the active tab, inject content.js, and invoke it.
  button.addEventListener("click", async () => {
    button.disabled = true;
    setStatus("Looking for active Teams tab...", "info");

    try {
      const [tab] = await chrome.tabs.query({
        active: true,
        currentWindow: true,
      });

      if (!tab || !tab.id) {
        setStatus("No active tab found.", "error");
        button.disabled = false;
        return;
      }

      if (!isTeamsUrl(tab.url)) {
        setStatus(
          "This tab doesn't look like a Microsoft Teams page. Open a Teams transcript and try again.",
          "error"
        );
        button.disabled = false;
        return;
      }

      setStatus("Scrolling and scraping transcript. This may take a moment...", "info");

      // Inject the content script into the page. Manifest V3 requires
      // scripting.executeScript rather than a declared content_script.
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ["content.js"],
      });

      // Kick off the scrape by messaging the content script.
      const response = await chrome.tabs.sendMessage(tab.id, {
        type: "MEETMARK_EXPORT",
      });

      if (!response || !response.ok) {
        const reason =
          (response && response.error) ||
          "Could not find transcript content. Make sure you're on a Teams transcript page and it has fully loaded.";
        setStatus(reason, "error");
        button.disabled = false;
        return;
      }

      // Hand the formatted markdown off to the service worker for download.
      const downloadResult = await chrome.runtime.sendMessage({
        type: "MEETMARK_DOWNLOAD",
        filename: response.filename,
        markdown: response.markdown,
      });

      if (downloadResult && downloadResult.ok) {
        if (response.warning) {
          setStatus("Exported with warning: " + response.warning, "warn");
        } else {
          setStatus("Transcript downloaded as " + response.filename, "success");
        }
      } else {
        setStatus(
          "Download failed: " +
            ((downloadResult && downloadResult.error) || "unknown error"),
          "error"
        );
      }
    } catch (err) {
      setStatus("Error: " + (err && err.message ? err.message : String(err)), "error");
    } finally {
      button.disabled = false;
    }
  });
});

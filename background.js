// MeetMark service worker.
//
// Handles download requests from the popup. Three payload shapes are
// supported:
//
//   1. { type: "MEETMARK_DOWNLOAD", filename, content, mime, format: "md" }
//      — write a Markdown file.
//   2. { type: "MEETMARK_DOWNLOAD", filename, content, mime, format: "txt" }
//      — write a plain text file.
//   3. { type: "MEETMARK_DOWNLOAD", filename, content, mime, format: "pdf" }
//      — open the printable HTML in a new tab. The HTML auto-invokes
//      window.print() once it loads, so the user lands in the browser's
//      print dialog and can Save as PDF. No actual .pdf file is written
//      directly — Chrome's native print-to-PDF does the conversion.
//   4. { type: "MEETMARK_DOWNLOAD", filename, dataUrl, mime: "image/png",
//        format: "png" }
//      — write a PNG file already encoded as a data: URL by the popup.
//
// In all disk-write cases we hand the data: URL to chrome.downloads.download.
// Data URLs work from an MV3 service worker where Blob URLs are not
// supported.

function utf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  return btoa(binary);
}

// Build a data URL for a UTF-8 text payload.
function buildTextDataUrl(content, mime) {
  const type = mime || "text/plain;charset=utf-8";
  return "data:" + type + ";base64," + utf8ToBase64(content);
}

function startDownload(url, filename, saveAs, sendResponse) {
  // Prepend subfolder if embedded in filename (e.g. "MeetMark/file.png").
  // chrome.downloads.download accepts relative paths within the Downloads dir.
  try {
    chrome.downloads.download(
      {
        url,
        filename,
        saveAs: !!saveAs,
      },
      (downloadId) => {
        if (chrome.runtime.lastError) {
          sendResponse({
            ok: false,
            error: chrome.runtime.lastError.message,
          });
          return;
        }
        if (typeof downloadId !== "number") {
          sendResponse({ ok: false, error: "Download did not start." });
          return;
        }
        sendResponse({ ok: true, downloadId });
      }
    );
  } catch (err) {
    sendResponse({
      ok: false,
      error: (err && err.message) || String(err),
    });
  }
}

// Open the printable HTML in a new tab. We pass the HTML as a data URL so
// the page runs at a sandbox-safe origin; it can still call window.print()
// against its own window, which is all we need. The page's inline script
// (injected by content.js's markdownToPrintableHtml) fires print() on load.
function openPrintTab(html, sendResponse) {
  try {
    const dataUrl = "data:text/html;charset=utf-8;base64," + utf8ToBase64(html);
    chrome.tabs.create({ url: dataUrl, active: true }, (tab) => {
      if (chrome.runtime.lastError) {
        sendResponse({
          ok: false,
          error: chrome.runtime.lastError.message,
        });
        return;
      }
      sendResponse({ ok: true, tabId: tab && tab.id });
    });
  } catch (err) {
    sendResponse({
      ok: false,
      error: (err && err.message) || String(err),
    });
  }
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!message || message.type !== "MEETMARK_DOWNLOAD") {
    return false;
  }

  try {
    const format = (message.format || "md").toString();
    const saveAs = !!message.saveAs;
    const filename = (message.filename || "transcript.md").toString();

    if (format === "pdf") {
      // The "content" is printable HTML — hand it to a new tab for the
      // user to Save as PDF from the print dialog.
      const html = (message.content || "").toString();
      openPrintTab(html, sendResponse);
      return true;
    }

    if (format === "png") {
      // The popup has already produced a full data: URL (data:image/png;
      // base64,...) from a canvas. Hand it straight to downloads.
      const url = (message.dataUrl || "").toString();
      if (!url) {
        sendResponse({ ok: false, error: "No PNG data URL supplied." });
        return true;
      }
      startDownload(url, filename, saveAs, sendResponse);
      return true;
    }

    // Default: text (md or txt).
    const content = (message.content || "").toString();
    const mime = (message.mime || "text/markdown;charset=utf-8").toString();
    const url = buildTextDataUrl(content, mime);
    startDownload(url, filename, saveAs, sendResponse);
  } catch (err) {
    sendResponse({
      ok: false,
      error: (err && err.message) || String(err),
    });
  }

  return true;
});

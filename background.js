// MeetMark service worker.
// Listens for download requests from the popup, encodes the markdown body
// into a data URL, and hands it to chrome.downloads.download. Using a data
// URL (rather than URL.createObjectURL) keeps things working from inside a
// service worker, where Blob URLs are not supported in MV3.

// Convert a UTF-8 string into a base64 string. Done by hand because btoa()
// only handles Latin-1.
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

// Build a data URL for a Markdown payload.
function buildDataUrl(markdown) {
  const b64 = utf8ToBase64(markdown);
  return "data:text/markdown;charset=utf-8;base64," + b64;
}

// Wire up the message listener. The popup posts a MEETMARK_DOWNLOAD message
// containing the filename and markdown body, and we reply with { ok } or
// { ok: false, error }.
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!message || message.type !== "MEETMARK_DOWNLOAD") {
    return false;
  }

  try {
    const filename = (message.filename || "transcript.md").toString();
    const markdown = (message.markdown || "").toString();
    const url = buildDataUrl(markdown);

    chrome.downloads.download(
      {
        url,
        filename,
        saveAs: false,
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
          sendResponse({
            ok: false,
            error: "Download did not start.",
          });
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

  // Returning true keeps the message channel open for the async response.
  return true;
});

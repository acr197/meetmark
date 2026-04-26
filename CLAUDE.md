# MeetMark — Project instructions

Chrome and Firefox extension. Converts Microsoft Teams meeting transcripts (SharePoint Stream) to Markdown, plain text, or PDF. Also captures full-page screenshots. Pure in-browser, no network calls.

## Stack

JavaScript, WebExtensions API (MV3), SharePoint Stream DOM, Manifest V3.

## Target browsers

Chrome (primary) and Firefox. Both must stay functional. Firefox uses `browser_specific_settings.gecko` in manifest.json with ID `meetmark@extension`.

## Architecture

- `manifest.json` — MV3 manifest, host_permissions for sharepoint.com
- `popup.html` / `popup.js` — Toolbar popup UI, format selection, cancel/progress wiring
- `content.js` — Injected into SharePoint Stream tabs; scrapes transcript DOM, scrolls, exports; communicates with popup over a long-lived Port named `"meetmark"`
- `background.js` — Service worker; handles `captureVisibleTab` calls for the full-page screenshot feature; stitches tiles
- `screenshot.js` — Injected for full-page capture; coordinates scroll positions and tile capture requests back to background

## Key behaviors

Popup connects to content script via `chrome.runtime.connect` (Port `"meetmark"`). Content script streams progress messages (`panel search`, `scroll %`, `cleaning`, `downloading`) back to popup. Popup can send `{ type: "cancel" }` to abort mid-scroll.

Speaker detection uses a two-pass innerText line parser (`scrapeTurnsByPattern`), not DOM selectors. Pass 1 builds a confirmed-speakers set; Pass 2 groups lines into turns. Five header cases handled: inline, merged, name-then-time, time-then-name, bare-confirmed-name.

Transcript scraper targets SharePoint Stream only (`*.sharepoint.com`). Full-page PNG works on any http/https page.

## Versioning

Follow global CLAUDE.md rules: bump `manifest.json` `version` on every commit. Minor fix = 0.0.1 bump, new feature = 0.1 bump.

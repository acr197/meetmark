# MeetMark
Chrome and Firefox extension that converts Microsoft Teams meeting transcripts to clean Markdown / plain text / PDF, and can also capture any page as a single full-page PNG.

[30s Demo](https://youtu.be/ZLr_SPXFyQY) (Sensitive data censored)

## What It Does
Clicking the toolbar icon opens a small menu with two paths:

- **Grab transcript as…** — picks the output format for the current Teams / SharePoint Stream transcript: Markdown (`.md`, default), plain text (`.txt`), or PDF (opens a printable view and jumps to the browser's Save-as-PDF dialog). Markdown is selected by default, so it's still effectively one click.
- **Capture full page as PNG** — works on any http/https page. MeetMark injects a content script that scrolls the page through a grid of viewport positions and, for each, calls `chrome.tabs.captureVisibleTab`. The popup stitches the captured tiles into one offscreen canvas (or several, for very long pages) and saves the result as a PNG via `chrome.downloads`. This is a faithful port of [mrcoles / GoFullPage](https://github.com/mrcoles/full-page-screen-capture-chrome-extension) adapted for MV3 (long-lived port messaging, `captured` ack before each scroll so captures stay under Chrome's rate limit).

All processing runs locally — no network calls, no external APIs, no telemetry.

## Stack
JavaScript, WebExtensions API (Chrome MV3 / Firefox MV3), SharePoint Stream DOM, Markdown

## How It Works
- Clicking the toolbar icon starts the export immediately — no secondary button click required
- Targets the inner scrollable transcript container in SharePoint Stream and scrolls incrementally to trigger lazy-loaded content without resetting scroll position
- Captures speaker name, timestamp, and speech content using DOM selectors anchored to element structure rather than fragile class names
- Groups consecutive utterances from the same speaker into a single block instead of splitting on every render segment
- Deduplicates captured turns by timestamp + speaker key to handle any scroll overlap
- Exports a single .md file with title, date, source URL, and full speaker-attributed transcript
- All processing runs entirely in the browser — no network calls, no external APIs, no telemetry

## Setup

### Chrome / Chromium
1. Go to `chrome://extensions`
2. Enable **Developer mode** (toggle, top-right)
3. Click **Load unpacked** and select this folder

### Firefox
1. Go to `about:debugging#/runtime/this-firefox`
2. Click **Load Temporary Add-on**
3. Select the `manifest.json` file inside this folder

For permanent Firefox installation, submit via [addons.mozilla.org](https://addons.mozilla.org).

## Usage

### Transcript export (MD / TXT / PDF)
1. Navigate to a Teams meeting recording in SharePoint Stream
2. Open the Transcript panel (click **Transcript** in the video sidebar if it isn't already visible)
3. Click the MeetMark toolbar icon
4. The format dropdown defaults to Markdown — just click **Grab**, or pick plain text or PDF first
5. PDF opens a printable view and launches the browser print dialog; choose **Save as PDF** to write the file

### Full-page screenshot (PNG)
1. Open the page you want to capture — this works on any http/https site, not just Teams
2. Click the MeetMark toolbar icon, then click **Capture full page as PNG**
3. The page will scroll from bottom to top as MeetMark captures each viewport. Let it run — the whole page is captured in a few seconds
4. MeetMark downloads one PNG covering the entire document. For extremely long pages (beyond ~30,000 device pixels on either axis) the output is split into multiple PNG tiles so it stays within the browser's canvas size limits

## Privacy
MeetMark runs entirely in your browser. It does not send transcript content to any server, does not call any external API, and has no analytics or telemetry. The only permission that touches the network is `downloads`, which is used solely to save the generated Markdown file to your local disk.

## Relevance
Solves a real workflow gap for teams using Microsoft 365 — meeting transcripts are locked in Stream with no clean export path. This makes them usable for knowledge management, AI assistants, and documentation pipelines.

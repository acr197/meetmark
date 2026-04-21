# MeetMark
Chrome and Firefox extension that converts Microsoft Teams meeting transcripts to clean Markdown / plain text / PDF, and can also capture any page as a single full-page PNG.

[30s Demo](https://youtu.be/ZLr_SPXFyQY) (Sensitive data censored)

## What It Does
Clicking the toolbar icon opens a small menu with three paths:

- **Quick Grab MD** — one click and the current Teams / SharePoint Stream transcript is exported as Markdown, exactly like earlier versions of MeetMark.
- **Grab as…** — the same transcript pipeline, but pick the output format: Markdown (`.md`), plain text (`.txt`), or PDF (opens a printable view and jumps to the browser's Save-as-PDF dialog).
- **Capture full page as PNG** — works on any page. The extension scrolls the current tab top to bottom, captures each viewport via `chrome.tabs.captureVisibleTab`, stitches the captures into a single tall PNG, and downloads it. Modeled after Peter Coles' [full-page-screen-capture-chrome-extension](https://github.com/mrcoles/full-page-screen-capture-chrome-extension).

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
4. Either click **Quick Grab MD** for the historical one-click Markdown export, or pick **Markdown / Plain text / PDF** from the **Grab as** dropdown and click **Grab**
5. PDF opens a printable view and launches the browser print dialog; choose **Save as PDF** to write the file

### Full-page screenshot (PNG)
1. Open the page you want to capture — this works on any site, not just Teams
2. Click the MeetMark toolbar icon
3. Click **Capture full page as PNG**. The extension briefly scrolls the page while capturing each viewport, then stitches the result into a single tall PNG and downloads it
4. If the stitched image exceeds the browser's maximum canvas size, MeetMark falls back to multiple tiled PNGs (`<name>-1.png`, `<name>-2.png`, …)

## Privacy
MeetMark runs entirely in your browser. It does not send transcript content to any server, does not call any external API, and has no analytics or telemetry. The only permission that touches the network is `downloads`, which is used solely to save the generated Markdown file to your local disk.

## Relevance
Solves a real workflow gap for teams using Microsoft 365 — meeting transcripts are locked in Stream with no clean export path. This makes them usable for knowledge management, AI assistants, and documentation pipelines.

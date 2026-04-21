# MeetMark
Chrome and Firefox extension that converts Microsoft Teams meeting transcripts to clean Markdown / plain text / PDF, and can also capture any page as a single full-page PNG.

[30s Demo](https://youtu.be/ZLr_SPXFyQY) (Sensitive data censored)

## What It Does
Clicking the toolbar icon opens a small menu with two paths:

- **Grab transcript as…** — picks the output format for the current Teams / SharePoint Stream transcript: Markdown (`.md`, default), plain text (`.txt`), or PDF (opens a printable view and jumps to the browser's Save-as-PDF dialog). Markdown is selected by default, so it's still effectively one click.
- **Capture full page as PNG** — works on any http/https page. MeetMark attaches to the tab via the Chrome DevTools Protocol (`chrome.debugger` + `Page.captureScreenshot({ captureBeyondViewport: true })`) and returns one PNG covering the entire document, including content rendered inside inner scrollable containers like dashboards or Fluent UI panels (cases where plain window scrolling does nothing). This is the same technique GoFullPage uses.

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
3. Chrome shows a yellow bar ("MeetMark started debugging this browser") for the few seconds the capture runs. That's normal — MeetMark needs the debugger API to render content that lives inside inner scrollable containers
4. MeetMark downloads a single PNG covering the entire document. Close DevTools before running if it complains about another debugger being attached

## Privacy
MeetMark runs entirely in your browser. It does not send transcript content to any server, does not call any external API, and has no analytics or telemetry. The only permission that touches the network is `downloads`, which is used solely to save the generated Markdown file to your local disk.

## Relevance
Solves a real workflow gap for teams using Microsoft 365 — meeting transcripts are locked in Stream with no clean export path. This makes them usable for knowledge management, AI assistants, and documentation pipelines.

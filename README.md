# MeetMark
Chrome and Firefox extension that converts Microsoft Teams meeting transcripts to clean Markdown files ready for AI ingestion.

[30s Demo](https://youtu.be/ZLr_SPXFyQY) (Sensitive data censored)

## What It Does
Opens directly on any Teams recording page in SharePoint Stream. Click the extension icon and the transcript is read immediately — no extra steps. A progress popup shows live status with the option to cancel or re-export. Output is a structured Markdown file with metadata, speaker blocks, and full dialogue, formatted for use with custom GPTs, RAG pipelines, or any LLM that ingests documents.

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
1. Navigate to a Teams meeting recording in SharePoint Stream
2. Open the Transcript panel (click **Transcript** in the video sidebar if it isn't already visible)
3. Click the MeetMark toolbar icon — export starts automatically
4. A progress popup appears; click **Cancel** to abort (closes the popup immediately), or wait for the download to complete
5. Click **Export again** to re-export without closing the popup

## Privacy
MeetMark runs entirely in your browser. It does not send transcript content to any server, does not call any external API, and has no analytics or telemetry. The only permission that touches the network is `downloads`, which is used solely to save the generated Markdown file to your local disk.

## Relevance
Solves a real workflow gap for teams using Microsoft 365 — meeting transcripts are locked in Stream with no clean export path. This makes them usable for knowledge management, AI assistants, and documentation pipelines.

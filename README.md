# MeetMark

A browser extension that exports Microsoft Teams meeting transcripts from SharePoint Stream to clean Markdown, plain text, or PDF. Also captures full-page screenshots of any webpage.

[Video demo](https://youtu.be/ZLr_SPXFyQY) (sensitive data censored)

## Features

- One-click export of Teams transcripts with speaker names, timestamps, and full attribution
- Output as Markdown (.md), plain text (.txt), or PDF via the browser's print dialog
- Groups consecutive lines from the same speaker into a single block instead of splitting on every render segment
- Deduplicates captured turns by timestamp and speaker to handle scroll overlap
- Full-page screenshot tool works on any webpage, not just Teams
- Stitches viewport tiles into a single PNG, splitting into multiple tiles for very long pages that exceed canvas limits
- Runs on both Chrome and Firefox

## Tech Stack

- **JavaScript** with the WebExtensions API (Manifest V3) for Chrome and Firefox
- **SharePoint Stream DOM** parsing for transcript extraction, using element structure rather than fragile class names
- **Offscreen Canvas** for tile stitching during full-page captures
- No external APIs, no network calls, no build step

## Privacy

- Everything runs in your browser. Transcript content never leaves your machine.
- No external API calls, no analytics, no telemetry.
- The only permission that touches the network is `downloads`, used solely to save the exported file to disk.

## Install

### Chrome
1. Go to `chrome://extensions` and enable Developer Mode
2. Click Load unpacked and select this folder

### Firefox
1. Go to `about:debugging` > This Firefox > Load Temporary Add-on
2. Select `manifest.json` from this folder

## Colophon

Published by AppCaddy.

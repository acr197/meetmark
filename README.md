# MeetMark
Chrome extension that converts Microsoft Teams meeting transcripts to clean Markdown files ready for Custom GPT ingestion for AI training.

## What It Does
Opens directly on any Teams recording page in SharePoint Stream. Scrapes the transcript panel, groups speech by speaker, and exports a structured Markdown file with metadata, speaker blocks, and full dialogue. Output is formatted for use with custom GPTs, RAG pipelines, or any LLM that ingests documents.

## Stack
JavaScript, Chrome Extensions API, SharePoint Stream DOM, Markdown

## How It Works
- Targets the inner scrollable transcript container in SharePoint Stream and scrolls incrementally to trigger lazy-loaded content without resetting scroll position
- Captures speaker name, timestamp, and speech content using DOM selectors anchored to element structure rather than fragile class names
- Groups consecutive utterances from the same speaker into a single block instead of splitting on every render segment
- Deduplicates captured turns by timestamp + speaker key to handle any scroll overlap
- Exports a single .md file with title, date, source URL, and full speaker-attributed transcript

## Setup
Load the unpacked extension in Chrome via `chrome://extensions`. Navigate to any Teams meeting recording in SharePoint Stream and open the Transcript panel. Click Export Transcript.

## Relevance
Solves a real workflow gap for teams using Microsoft 365 -- meeting transcripts are locked in Stream with no clean export path. This makes them usable for knowledge management, AI assistants, and documentation pipelines.

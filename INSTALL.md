# Installing MeetMark

MeetMark is a Chrome extension that exports Microsoft Teams meeting transcripts
to clean Markdown files. It is not published to the Chrome Web Store — you load
it as an unpacked extension in developer mode.

## Requirements

- Google Chrome (or any Chromium-based browser: Edge, Brave, Arc, etc.)
- Access to a Microsoft Teams meeting transcript page in your browser

## Steps

1. **Download the source.** Clone or download this repository to a folder on
   your machine. Note the full path — Chrome will need it.

2. **Open the extensions page.** In Chrome, navigate to:

   ```
   chrome://extensions
   ```

3. **Enable Developer mode.** Toggle the "Developer mode" switch in the upper
   right corner of the extensions page.

4. **Load unpacked.** Click the "Load unpacked" button that appears, then
   select the folder containing this extension (the folder with `manifest.json`
   directly inside it).

5. **Pin the extension (optional).** Click the puzzle-piece icon in the Chrome
   toolbar, find "MeetMark", and click the pin icon so the button stays
   visible.

## Using it

1. Open a Microsoft Teams meeting recording in your browser. For most
   tenants this is a SharePoint Stream page with a URL like
   `https://<tenant>-my.sharepoint.com/.../stream.aspx?id=...`.
2. Click the **Transcript** button on the right side of the player so the
   transcript panel is visible. Wait for it to render.
3. Click the MeetMark toolbar button.
4. Click **Export transcript**.
5. The extension will scroll the transcript panel to load all content,
   scrape it, clean it up, and download a `.md` file named
   `YYYY.MM.DD - {Meeting Title} Transcript.md` to your default downloads
   folder.

## Troubleshooting

- **"Could not find transcript content."** Make sure you're actually on a
  Teams transcript page and that it has fully rendered before clicking export.
  Teams sometimes shows a loading spinner for several seconds.
- **"Scroll timed out."** The export still completes with whatever was
  loaded; very long meetings may need a manual scroll first.
- **Nothing happens when clicking the button.** Open the popup, right-click,
  choose "Inspect", and check the console for errors. The content script also
  logs which selectors matched in the page console (F12 on the Teams tab).
- **Teams updated their UI.** MeetMark uses fallback selectors and logs which
  one matched. If none match, the selector lists in `content.js` need to be
  updated.

## Updating

After editing any file in this folder, return to `chrome://extensions` and
click the circular reload icon on the MeetMark card. Then reload your Teams
tab so the new content script is injected.

## Privacy

MeetMark runs entirely in your browser. It does not send transcript content
to any server, does not call any API, and has no analytics or telemetry. The
only network access in the manifest is to teams.microsoft.com itself, which
is the page being scraped.

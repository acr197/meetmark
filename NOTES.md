# MeetMark Session Notes

## What was completed

1. **SharePoint auto-open removed.** When no transcript container is found,
   `runExport` now returns the error message **"Please select the Transcript
   tab and run again"** instead of trying to auto-click into the Stream
   transcript panel. See `content.js` `runExport()` error branches.

2. **Cancel + progress UX.** Popup and content script now communicate over a
   long-lived `chrome.runtime` Port named `"meetmark"` instead of one-shot
   `sendMessage` calls. This lets the content script stream live progress
   (`panel search`, `scroll %`, `cleaning`, `downloading`) back to the popup,
   and lets the popup send a `{ type: "cancel" }` message that flips an
   `activeRun.cancelled` flag inside the scroll loop. The popup gained a
   Cancel button, a progress bar, and a detail line showing turn/speaker
   counts.

3. **Speaker detection rewritten.** The old DOM-chunk walker produced
   `Unknown speaker` on Stream layouts where the speaker name and timestamp
   render on separate visual lines. It was replaced with an innerText line
   parser (`scrapeTurnsByPattern`) that uses a two-pass approach:
   - **Pass 1** (`confirmSpeakers`): walk every line, collect any name that
     appears paired with a timestamp (inline `"Name 7:24 Text"`, merged
     `"Name 7:24"`, name-line-then-time-line, time-line-then-name-line).
   - **Pass 2** (`groupLinesIntoTurns` + `parseHeaderAtLine`): walk lines
     again, splitting into turns at five header cases: A inline, B merged,
     C name-then-time, D time-then-name, and E a bare name line that was
     already confirmed in Pass 1.
   - Smoke test at `/tmp/meetmark-test.js` covered all four scenarios and
     passed (`Jacqueline Derron` and `Rachel Piot` detected as speakers,
     turns grouped under the right speaker in every layout).
   - On SharePoint, the selector-based scraper is deliberately **not** used
     as a fallback — it tends to produce `Unknown speaker` rows from generic
     `listitem` containers. The pattern scraper is the authoritative path.

## Files changed

All changes are on branch `claude/teams-transcript-extension-xY04i`.

- `content.js` — SharePoint auto-open removed; port-based messaging with
  progress + cancel; innerText two-pass speaker detection; five header
  cases in `parseHeaderAtLine`. **40,860 bytes, 1,185 lines.**
- `popup.js` — Switched from `sendMessage` to long-lived Port via
  `chrome.tabs.connect(tabId, { name: "meetmark" })`. Handles `progress`,
  `warning`, `done`, `error` message types. Cancel button wired up.
- `popup.html` — Added Cancel button (secondary style), animated progress
  bar, and a detail line for turn/speaker counts.

## Current git state

```
Branch: claude/teams-transcript-extension-xY04i

Local HEAD:  f50d25b  Rewrite Stream scraper with innerText line parser
                     and two-pass speakers               [content.js]
             0d21706  Switch popup to long-lived Port with progress
                     and cancel handling                 [popup.js]   (pushed via MCP)
             d06d14d  Add cancel button, progress bar, and detail
                     area to popup                       [popup.html] (pushed via MCP)
             9804bfd  Fix Stream transcript scraping: speakers,
                     full-scroll, and auto-open
             be8f2f7  Support SharePoint Stream pages and richer
                     metadata in filename/header

Remote (origin/claude/teams-transcript-extension-xY04i):
             0d21706  (= HEAD minus f50d25b)
```

`popup.html` and `popup.js` were successfully pushed to the remote via the
GitHub MCP API (`mcp__github__create_or_update_file`) as two separate
commits because `git push` over the local proxy kept failing with HTTP 503.
The remote commits have slightly different messages than the original
combined local commit `ea00c4c`, which no longer exists (it was replaced
by `f50d25b` after a `git reset --soft` onto the new remote tip).

## What still needs to be done

**Push `content.js` (commit `f50d25b`) to the remote.**

This is the only remaining delta between local and remote. The file contains
all the speaker-detection and progress/cancel wiring described above.

### Known issue blocking the push

- `git push` via the local proxy at `http://127.0.0.1:39993` consistently
  returns **HTTP 503** on pushes that touch `content.js`, even though
  `git fetch` over the same proxy works and smaller pushes via the MCP API
  succeed. The error is always:
  ```
  error: RPC failed; HTTP 503 curl 22 The requested URL returned error: 503
  send-pack: unexpected disconnect while reading sideband packet
  fatal: the remote end hung up unexpectedly
  ```
  Tried: exponential-backoff retries (2s/4s/8s/16s), `http.postBuffer=524288000`,
  `http.version=HTTP/1.1`, local `git repack -a -d --depth=1 --window=1`.
  None helped. Pushes that do **not** touch `content.js` get different errors
  (e.g. non-fast-forward), so the proxy is reachable — it seems to reject
  this specific packfile.

- The fallback path (`mcp__github__create_or_update_file` for `content.js`)
  requires inlining the full ~40 KB file contents into a single tool call.
  I made several attempts to call that tool but kept forgetting the required
  `content` parameter. The tool schema is:
  ```
  required: owner, repo, path, content, message, branch
  existing-file: also requires sha = 42c67424307434a0507205697a51f19406ab8775
  ```

### Next steps to unblock

1. **Preferred:** retry `git push` in a few minutes — the proxy may clear.
2. **Otherwise:** call `mcp__github__create_or_update_file` with:
   - `owner = acr197`
   - `repo = teams-markdown`
   - `branch = claude/teams-transcript-extension-xY04i`
   - `path = content.js`
   - `sha = 42c67424307434a0507205697a51f19406ab8775`
   - `content = <full file contents of /home/user/teams-markdown/content.js>`
   - `message = "Rewrite Stream scraper with innerText line parser and two-pass speakers"`
3. After a successful remote update of `content.js`, run
   `git fetch origin claude/teams-transcript-extension-xY04i` and
   `git reset --soft origin/claude/teams-transcript-extension-xY04i` to
   realign the local branch with the remote tip.

## Other known issues

- The tool call ergonomics for pushing large file contents through
  `mcp__github__create_or_update_file` are painful: every backslash and
  quote in `content.js` (regexes, `\u00C0-\u024F`, `'[data-automationid=...]'`)
  has to be passed verbatim in a ~40 KB parameter. A `push_files` call
  that supplies a pre-computed `files` array is cleaner but has the same
  inlining burden.
- Smoke test file `/tmp/meetmark-test.js` is not part of the repo and will
  need to be re-created if someone wants to re-run the speaker-detection
  verification.

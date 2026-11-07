# Privacy and network use

This file lists every way the code in this repository can send data off the reader's device,
and everything it keeps on the device. It covers the published packages (`ooxml-core`,
`ooxml-ui` and every package under `viewers/`), the OOXML Office suite (`apps/office-suite`,
`site/`), the desktop shell (`desktop/`), the MCP servers (`mcp/`) and the documentation sites.

The libraries are local-first. Opening, rendering, editing and saving a document happens in the
browser (or in Node) and never uploads the file. Nothing here collects analytics or telemetry,
reports errors to a service, or checks for updates against a server of ours. There is no server
of ours: every endpoint below is either a fixed third party named here, or one that the host
application or the user configures.

When a statement here and the code disagree, the code is wrong or this file is out of date:
please [open an issue](https://github.com/ChristopherVR/ooxml/issues). A change that adds a
network request must update this file in the same commit.

## Summary

| What                                             | Goes to                                     | On by default                     | Turn it off or redirect it                                         |
| ------------------------------------------------ | ------------------------------------------- | --------------------------------- | ------------------------------------------------------------------ |
| PowerPoint: fonts not installed or embedded      | `fonts.googleapis.com`, `fonts.gstatic.com` | Yes                               | `googleFonts` option; File > Options > Trust Center; suite setting |
| Word, Excel, Visio: fonts not installed          | `fonts.googleapis.com`, `fonts.gstatic.com` | No (yes in the suite)             | `googleFonts` property; Trust Center (off by default); suite       |
| PowerPoint: linked pictures, media, video embeds | Any URL the deck links to                   | Yes in the viewer, no in the core | `allowExternalContent` option; Trust Center; suite setting         |
| PowerPoint: Share / Broadcast without a server   | `wss://y-webrtc-eu.fly.dev`, public STUN    | Only when the user starts it      | Configure a signaling server or `serverUrl`                        |
| PowerPoint: Share / Broadcast with a server      | The host's or user's server                 | Only when the user starts it      | Do not start a session                                             |
| PowerPoint (Node): signature inspection          | The OCSP responder named in the certificate | Yes, when inspection is called    | No switch today (see below)                                        |
| Teams: sync and signaling                        | The configured `syncUrl` and `signalingUrl` | Only when a server is configured  | Local mode                                                         |
| Teams: calls                                     | `stun:stun.l.google.com:19302`, then peers  | Yes, when a call starts           | `iceServers` (your own STUN/TURN, or `[]`)                         |
| Teams: website tabs                              | The URL a member added                      | When the tab is opened            | Do not add or open website tabs                                    |
| Suite: AI assistant                              | The provider the user or policy selects     | No (offline by default)           | `ai-config.json`                                                   |
| Suite: OneDrive                                  | `graph.microsoft.com`, Microsoft sign-in    | No (host must enable it)          | Do not configure it                                                |
| Suite: app updates                               | The suite's own origin (`sw.js`)            | Yes                               | Not applicable (same origin)                                       |
| Any product: hyperlinks, help links              | The link target                             | Only when clicked                 | Not applicable                                                     |
| Host-supplied document URLs                      | The URL the host passes                     | Only when the host passes one     | Pass bytes instead                                                 |

Every request also sends what any browser request sends: the reader's IP address, the browser's
user agent and, for some requests, a `Referer`. That is the main privacy cost of each third-party
destination above, even when no document content is sent.

## PowerPoint (`ooxml-ui/pptx` and the `pptx-*-viewer` packages)

### Fonts that are neither installed nor embedded

A deck often names fonts the reader does not have: Microsoft 365 cloud fonts, theme fonts such
as Aptos, or a template's display face. PowerPoint downloads those from Microsoft on demand. The
viewer does the same from Google Fonts, so the text keeps its intended look and line breaks.

- **Trigger:** opening a deck that uses such a font, when the bundled Google Fonts catalogue
  serves that family or a metric-compatible clone of it (for example Carlito for Calibri).
- **Sent:** a request to `https://fonts.googleapis.com/css2` whose query names the families
  (and weights) the deck needs, then requests to `fonts.gstatic.com` for the font files. To save
  time, a first request is made while the deck is still opening, naming the catalogue families
  its theme, masters and first slides declare, so it can include a family the deck declares but
  does not draw; font files are only downloaded for faces actually drawn.
  Families that are installed on the device, or embedded in the deck, are never named. The
  viewer never probes Google for families its catalogue does not list, so the names of private
  or corporate fonts are not sent.
- **Also covered by the same setting:** the WordArt glyph-outline download (only for warped
  WordArt in such a font) and the fonts inlined into slide image exports.
- **Off or redirected:**
  - the host option `googleFonts: false` turns it off; `googleFonts: { baseUrl }` sends the
    same requests to a self-hosted mirror instead;
  - the reader can turn it off in File > Options > Trust Center (the host option, when set,
    locks that toggle);
  - in the OOXML Office suite, Settings > Fonts.

With downloads off, text uses installed, embedded or substitute fonts, so line breaks can differ
from PowerPoint. [docs/font-loading.md](docs/font-loading.md) describes the whole font pipeline.

### Linked external content

A deck can link (rather than embed) pictures, shape fills, picture bullets, 3D model posters,
audio and video at an external URL, and can insert an online video (YouTube, Vimeo) that is
shown in the provider's embed frame.

- **Trigger:** opening a deck with such links.
- **Sent:** requests to each linked URL, which tells that server that the deck was opened.
  An online video frame loads the provider's page and whatever it loads.
- **Default:** the viewer's Trust Center option "Allow external content" is on. The core
  parser's own `allowExternalImages` option is off, so `ooxml-core/pptx` used directly fetches
  nothing.
- **Off:** the host option `allowExternalContent: false`, or the Trust Center option. Teams
  presentation previews always block it. In the OOXML Office suite, Settings > Privacy >
  "Block linked pictures and media in presentations" passes `false` to every deck opened next.

### Collaboration: Share and Broadcast

Sessions start only when the user (or host) starts one.

- With a server URL, document updates, presence and the room name go to that WebSocket server.
- With the server field left blank, the viewer uses serverless WebRTC (`y-webrtc`). Peers find
  each other through the library's default public signaling server, `wss://y-webrtc-eu.fly.dev`,
  operated by a third party, and through public STUN servers (`stun.l.google.com`,
  `global.stun.twilio.com`) to learn their network addresses. The signaling server sees the
  room id (hashed), the participants' IP addresses and the connection offers; document content
  then flows directly between peers, encrypted by WebRTC. Pass a `signaling` list (or
  `?signaling=` in a share link) to use your own server.
- Sessions within one browser use `BroadcastChannel` and leave the device.

### Digital signature inspection (Node only)

`inspectPptxDigitalSignatures` from `ooxml-core/pptx/signature-node` checks each signing
certificate's revocation status by sending an OCSP request (the certificate's serial number and
issuer hashes) to the responder URL written in that certificate. Requests to private network
addresses are refused. There is currently no option to skip the check; call the browser-safe
signature detection in `ooxml-core/opc` instead when no network use is acceptable.

## Word, Excel and Visio

By default the Word, Excel and Visio packages make no network request when opening, rendering,
editing or saving a document:

- linked (external) pictures are not fetched;
- external workbook links, Visio data record sets and linked chart data are kept but never
  refreshed; Excel web queries and `WEBSERVICE` are not evaluated;
- fonts come from the device and the file only, unless the host sets `googleFonts` (below);
- workers are loaded from the page's own origin.

**Fonts, when the host or the reader turns them on.** `<docx-editor>`, `<xlsx-editor>` and
`<visio-viewer>` take the same `googleFonts` value as PowerPoint, but unset leaves it to the
reader's File > Options > Trust Center toggle "Download fonts that are not installed or embedded",
which is off by default (a host value locks the toggle). With
`{ baseUrl: 'https://fonts.googleapis.com/css2' }` (or a self-hosted mirror) they send the same
requests PowerPoint does, from the same shared code (`ooxml-core/text`, `ooxml-ui/fonts`): the
public catalogue families (or metric-compatible clones such as Carlito for Calibri) the file
declares (theme, styles, font table) and the device lacks, then the font files for the faces it
draws. The file, its text and the names of fonts outside
Google's public catalogue are never sent. `false` or an invalid `baseUrl` sends nothing.

Hyperlinks and in-product help links open only when the reader follows them.

## Teams (`ooxml-core/teams`, `ooxml-ui/teams`, the `openteams-*` packages)

- **Sync and signaling:** chat, presence and session data go only to the `syncUrl` and
  `signalingUrl` the host or the user configures. Local mode stays in the browser
  (`BroadcastChannel`). The reference server in `viewers/teams/server` makes no outbound calls.
- **Calls:** without an `iceServers` setting, calls use Google's public STUN server
  `stun:stun.l.google.com:19302`, which learns each participant's IP address. Set `iceServers`
  to your own STUN/TURN servers, or to `[]` to use none. Audio and video flow between peers or
  through the TURN server you configure.
- **Website tabs:** a website tab loads the URL a member added, in a frame, for everyone who
  opens it.

## OOXML Office suite (`apps/office-suite`, the GitHub Pages site and the desktop app)

The suite uses the editors above, so their sections apply. In addition:

- **Fonts:** Settings > Fonts chooses Google Fonts (the default), your own font server, or no
  downloads, for PowerPoint, Word, Excel and Visio alike. The site's Content Security Policy allows only Google's two font
  hosts; a deployment that uses its own font server adds that server to the policy.
- **AI assistant:** off the network by default (the shipped `ai-config.json` selects the offline
  assistant). When the user or the deployment's policy selects a provider, the prompt, the
  conversation and the document context it needs are sent to that provider, under the
  provider's own terms. Keys stay in session storage unless the user chooses to remember them.
  See [docs/suite-ai-config.md](docs/suite-ai-config.md).
- **OneDrive:** off unless the host configures it. Files and the access token then go to
  Microsoft Graph. See [docs/onedrive-integration.md](docs/onedrive-integration.md).
- **Shared AI chats:** sync over the configured collaboration or Teams server, otherwise stay in
  this browser.
- **Updates:** the installed app checks its own origin for a new `sw.js` at start, on focus and
  hourly. The service worker caches only the app's own files.
- **Desktop:** the Tauri shell has no updater, telemetry or plugins and reads and writes only
  the files the user opens. The operating system's WebView (WebView2, WKWebView) may send its
  vendor's own diagnostics under the operating system's settings.

## Data kept on the device

Nothing below is sent anywhere by this code. Clearing the site's data removes it.

- **PowerPoint:** autosave and recovery snapshots, the session, presenter and audience state in
  IndexedDB (`pptx-viewer-autosave`, `pptx-viewer-session`, `pptx-viewer-presentation`,
  `pptx-viewer-audience`); preferences, including Trust Center choices, in localStorage
  (`pptx-viewer-prefs`). File > Account > Storage & Privacy shows and clears them. Fonts the
  user adds under File > Options > Custom fonts are kept for the session only.
- **Word, Excel, Visio:** small preferences in localStorage (view options, custom lists, saved
  shapes), including the Trust Center font-download choice (`ooxml.docx.allowFontDownloads`,
  `ooxml.xlsx.allowFontDownloads`, and `allowFontDownloads` in Visio's `ooxml-visio-options`).
- **Teams:** chat snapshots, read state, drafts, presence and follows in the storage the host
  provides; the server configuration in localStorage. Session tokens are kept in memory only.
- **Suite:** the document library, attachments, versions and chat history in IndexedDB per
  profile; profiles, favourites, theme (`ooxml-suite-theme`), fonts (`ooxml-suite-fonts`), the
  other Settings (`ooxml-suite-settings`) and AI preferences in localStorage; AI keys in session
  storage, or localStorage when the user chooses to remember them. Settings > Privacy shows how
  much space the app uses and can ask the browser to keep it when space runs low
  (`navigator.storage.persist`); nothing is sent.

## Browser features outside our control

- Text boxes enable the browser's spell checker. Some browsers' enhanced spell checking sends
  text to the browser vendor; that is a browser setting.
- Browser extensions can read any page, including open documents.

## Documentation sites

The documentation sites (`christophervr.github.io/ooxml/...`) bundle their own fonts and search
index and make no third-party requests. Pages that embed a demo apply the sections above.

## Development tooling

Scripts that are not shipped contact external services when a maintainer runs them: the npm
registry, GitHub, Discord announcements, Google Fonts' metadata feed (to refresh the font
catalogue) and a translation API (to draft locale files). None of them run in a reader's
browser.

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

| What                                               | Goes to                                     | On by default                            | Turn it off or redirect it                                         |
| -------------------------------------------------- | ------------------------------------------- | ---------------------------------------- | ------------------------------------------------------------------ |
| PowerPoint: fonts not installed or embedded        | `fonts.googleapis.com`, `fonts.gstatic.com` | Yes                                      | `googleFonts` option; File > Options > Trust Center; suite setting |
| Word, Excel, Visio: fonts not installed            | `fonts.googleapis.com`, `fonts.gstatic.com` | No (yes in the suite)                    | `googleFonts` property; Trust Center (off by default); suite       |
| PowerPoint: linked pictures, media, sounds, embeds | Any URL the deck links to                   | No: only after Enable Content            | `allowExternalContent` option; Trust Center; suite setting         |
| PowerPoint: Share / Broadcast without a server     | `wss://y-webrtc-eu.fly.dev`, public STUN    | Only when the user starts it             | Configure a signaling server or `serverUrl`                        |
| PowerPoint: Share / Broadcast with a server        | The host's or user's server                 | Only when the user starts it             | Do not start a session                                             |
| PowerPoint (Node): signature inspection            | The OCSP responder named in the certificate | When a signature value verifies          | `checkRevocation: false` (see below)                               |
| PowerPoint: AI assistant                           | The connection the host configures          | No (host must enable it)                 | Do not pass `ai`                                                   |
| Teams: sync and signaling                          | The configured `syncUrl` and `signalingUrl` | Only when a server is configured         | Local mode                                                         |
| Teams: calls                                       | `stun:stun.l.google.com:19302`, then peers  | Yes, when a call starts                  | `iceServers` (your own STUN/TURN, or `[]`)                         |
| Teams: website tabs                                | The URL a member added                      | When the tab is opened                   | Do not add or open website tabs                                    |
| Suite: AI assistant                                | The provider the user or policy selects     | No (offline by default)                  | `ai-config.json`                                                   |
| Suite: shared AI chat credentials                  | The policy's `collaboration.credentialsUrl` | No (host must configure it)              | Leave `collaboration.credentialsUrl` empty                         |
| Suite: OneDrive                                    | `graph.microsoft.com`, Microsoft sign-in    | No (host must enable it)                 | Do not configure it                                                |
| Suite: app updates                                 | The suite's own origin (`sw.js`)            | Yes                                      | Not applicable (same origin)                                       |
| Any product: hyperlinks, help links                | The link target, in a new tab               | Only when clicked                        | Not applicable                                                     |
| Word, Excel, PowerPoint: Get Add-ins               | `marketplace.microsoft.com`                 | Only when its link is clicked            | Leave the catalog link unopened                                    |
| Excel, PowerPoint: Forms                           | `forms.office.com`, `forms.cloud.microsoft` | Only when the user opens or loads a form | Leave the form unopened or close the response dialog               |
| Excel: `WEBSERVICE` and `IMAGE` fetch providers    | Any `http:`/`https:` URL in the formula     | No (host must opt in via `providers`)    | Omit `providers.webservice`/`image` or restrict `allowedOrigins`   |
| Host-supplied document URLs                        | The URL the host passes                     | Only when the host passes one            | Pass bytes instead                                                 |

Every request also sends what any browser request sends: the reader's IP address, the browser's
user agent and, for some requests, a `Referer`. That is the main privacy cost of each third-party
destination above, even when no document content is sent.

## Screenshots and screen recordings

- **Trigger:** selecting Screenshot or Screen Clipping in PowerPoint, Word or Excel, or
  starting a Screen Recording in PowerPoint. The browser asks the user which surface to share
  each time. For a recording, the user can also share supported surface audio, request pointer
  inclusion or exclusion, and crop to an area of the already-shared surface in memory.
- **Data:** the selected screen, window or tab is processed in memory. A screenshot or recording
  is embedded in the open document only when inserted. The capture stops when the user finishes
  or cancels and its permission is never stored by the app.
- **Network:** capture makes no network requests. If the user inserts the result into a shared
  document, the existing collaboration settings govern transfer to peers. Photo Album and Reuse
  Slides read the chosen files locally. Existing linked-content rules apply to links in reused slides.

## Local 3D models and Office objects

Insert > 3D Models and Insert > Object in PowerPoint, Word and Excel read only the selected local
files. The editor stores the chosen file and a generated image preview in the document when
Insert completes. Local insertion makes no network request.

PowerPoint shares the inserted file and preview through Share or Broadcast when a session is
active, subject to the room's asset limits below. Word's opt-in collaboration also sends the
inserted bytes and preview to the host-configured room transport as document assets. Leave the
room or omit the collaboration configuration to keep insertions local.

## PowerPoint (`ooxml-ui/pptx` and the `pptx-*-viewer` packages)

Opening a multi-slide presentation can load the preview renderer from the origin that serves
the application's JavaScript. The request carries the asset URL and standard browser request
metadata, with no document data or filename. Hosts can bundle the renderer with their app to
avoid a separate request.

Picture Format reads the selected local PNG or JPEG, or an embedded picture's
local object URL, when a picture editing command is invoked. The browser loads
the picture worker from the application and processes the bytes on the device.
The worker receives no remote picture URLs. Keep Changes stores the edited
bitmap, original bitmap and marks in the presentation through the existing
save path. When AutoSave or AutoRecover stores a presentation snapshot, that
snapshot can include the picture edits. The existing AutoSave and AutoRecover
settings control those snapshots.
During an active PowerPoint collaboration session, a picture edit also shares
the edited image, its original bitmap and removal marks with the room through
the existing collaboration transport. Leave the room to keep edits local.

### Cameo camera preview

The Cameo button on a live slide surface can activate the device camera after the user
chooses it and the browser grants permission. Video is displayed locally and the track
stops when the preview is turned off or removed. The slide file contains a stock poster and
the live-feed descriptor.

Record > Cameo > Record camera clip requests camera access when the user selects Record.
Microphone access requires selecting Include microphone. Recording happens locally and the
clip becomes embedded video in the presentation if the user selects Stop and insert. Closing
the dialog stops capture and discards the unsaved clip. Once inserted, the video follows the
same save, share and collaboration behavior as other embedded media.
Reset to Cameo removes editor-recorded clips from the slide. The package writer can retain
the unused media parts when saving an existing file.

### Forms and add-in links

Get Add-ins in Word, Excel and PowerPoint opens a dialog with a link to Microsoft's
Marketplace. Forms in Excel and PowerPoint offers a link to the Microsoft Forms site and
inserts a Microsoft Forms response URL supplied by the user. Opening the dialog and entering
a link make no request. During a PowerPoint slide show, a Forms response link offers to load
the form in an isolated frame or open it in the browser. Loading the form sends its response
URL and normal request metadata, without a referrer. Answers entered in the frame go directly
to Microsoft. Closing the dialog removes the frame. Other link actions open the chosen
destination in a new tab. My Add-ins reads the host's local ribbon tabs.

Opening Get Add-ins or My Add-ins can load bundled dialog and manifest-reader code from the
app's origin. The request carries no selected-file or document data. Local installs start
when the user chooses an XML, JSON or ZIP file and confirms the preview. The editor keeps
the normalized manifest and SHA-256 digest in memory for that editor instance. Removing the
add-in or closing the editor clears its record. The app saves no install data to device storage
or the document and makes no requests to the publishers, asset origins or schema URLs in the
manifest. Local installs grant no document access.

### Fonts that are neither installed nor embedded

A deck often names fonts the reader does not have: Microsoft 365 cloud fonts, theme fonts such
as Aptos, or a template's display face. PowerPoint downloads those from Microsoft on demand. The
viewer does the same from Google Fonts, so the text keeps its intended look and line breaks.

- **Trigger:** opening a deck that uses such a font, when the bundled Google Fonts catalogue
  serves that family or a stand-in for it: a metric-compatible clone (for example Carlito for
  Calibri) or a measured look-alike whose widths the viewer corrects to the deck font's own
  (Aptos and Aptos Display as Gantari, Aptos Mono as Sometype Mono, Consolas as Overpass Mono,
  Comic Sans MS as Comic Relief, Garamond as EB Garamond, Impact as Anton).
- **Sent:** a request to `https://fonts.googleapis.com/css2` whose query names the families
  (and weights) the deck needs, then requests to `fonts.gstatic.com` for the font files. To save
  time, a first request is made while the deck is still opening, naming the catalogue families
  its theme, masters and first slides declare, so it can include a family the deck declares but
  does not draw; font files are only downloaded for faces actually drawn.
  Families that are installed on the device, or embedded in the deck, are never named. The
  viewer never probes Google for families its catalogue does not list, so the names of private
  or corporate fonts are not sent.
- **Also covered by the same setting:** the WordArt glyph-outline download (only for warped
  WordArt in such a font) and the fonts inlined into slide image exports (including the effects
  a vector PDF export draws as pictures). The vector PDF export itself requests nothing: it
  embeds only fonts the page already holds (the deck's embedded fonts, from `data:` or `blob:`
  URLs) and fonts a host passes to `exportSlidesVectorPdf`.
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
audio, video and the sounds that play with animations, slide transitions and hyperlinks at an
external URL, and can insert an online video (YouTube, Vimeo) that is shown in the provider's
embed frame.

- **Trigger:** choosing **Enable Content** on the security bar of a deck with such links, as in
  Microsoft PowerPoint, or a host or Trust Center setting that allows them without asking.
- **Sent:** requests to each linked URL, which tells that server that the deck was opened.
  A linked sound is requested when it plays in the slide show. An online video frame loads the
  provider's page and whatever it loads.
- **Default:** off. Opening a deck requests none of its linked content; a security bar ("References
  to external pictures and other media have been blocked") offers Enable Content for that
  presentation. Enabling lasts until another presentation is opened and is not stored. In a
  collaboration session the linked addresses travel with the shared deck, but each participant's
  viewer requests them only after that participant chooses Enable Content; one participant's
  choice is never sent to the others. The core
  parser's own `allowExternalImages` option is off too, so `ooxml-core/pptx` used directly
  fetches nothing. [docs/powerpoint-external-content.md](docs/powerpoint-external-content.md)
  records PowerPoint's own default and how it was established.
- **Hyperlinks:** a link opens only when clicked, always in a new tab with `noopener` and
  `noreferrer`, whatever frame the deck names. `file:` and `ftp:` links are kept in the deck but
  never opened.
- **On without asking:** the host option `allowExternalContent: true` (PowerPoint's "Unblock
  automatic download of linked images" policy), or File > Options > Trust Center > "Allow
  external content", which is off by default.
- **Off without the bar:** the host option `allowExternalContent: false`. Teams presentation
  previews always block it. In the OOXML Office suite, Settings > Privacy > "Block linked
  pictures and media in presentations" passes `false` to every deck opened next.

### Collaboration: Share and Broadcast

Sessions start only when the user (or host) starts one.

- With a server URL, document updates, presence and the room name go to that WebSocket server.
- With the server field left blank, the viewer uses serverless WebRTC (`y-webrtc`). Peers find
  each other through the library's default public signaling server, `wss://y-webrtc-eu.fly.dev`,
  operated by a third party, and through public STUN servers (`stun.l.google.com`,
  `global.stun.twilio.com`) to learn their network addresses. Each session has a random 128-bit
  secret that its invitation link carries in the `#key=` fragment, which browsers never send to
  a server. The signaling server sees a SHA-256 hash of that secret (never the room name), the
  participants' IP addresses and signaling messages encrypted with the secret; document content
  then flows directly between peers, encrypted by WebRTC and by the same secret. Pass a
  `signaling` list (or `?signaling=` in a share link) to use your own server.
- Sessions within one browser use `BroadcastChannel`, which never leaves the device.
- **What the room carries:** the slides, and the bytes of the pictures, poster frames, audio,
  video, 3D models and Office objects embedded in them, so a participant who has no copy of the
  file sees the same deck.
  Each participant turns those bytes into its own `blob:` or `data:` URLs; no URL of one page is
  sent to another. A room holds at most 16 MiB of such files (128 KiB in a peer-to-peer room,
  whose first message to a joiner must fit a WebRTC data channel); a larger one is left out and the
  other participants are told it was too large to share. The media reader accepts only browser-local
  `blob:` URLs. Linked (external) pictures and media
  are never fetched and re-sent: the room carries only their link, and each participant's own
  Enable Content choice (or `allowExternalContent`) decides whether its browser requests it.

### AI assistant

The assistant is off unless the host passes an `ai` configuration. When it is on, each request goes
to the connection the host chose (its endpoint, its transport, or a model running in the page) and
carries the conversation, a deck outline or the active slide (`contextStrategy`, `'none'` turns it
off), whatever the user focused, and the results of the tools the model called. This applies in every
framework binding. `tools.enabled` limits what the model can read. Stored chats stay on the device.

### Digital signature inspection (Node only)

`inspectPptxDigitalSignatures` from `ooxml-core/pptx/signature-node` checks the revocation status
of each certificate that verified a signature value, by sending an OCSP request (the certificate's
serial number and issuer hashes) to the responder URL written in that certificate. A signature
that fails verification causes no request. The URL comes from the file, so the request goes only
to globally routable addresses, and the connection is made to the address that was checked. Pass
`{ checkRevocation: false }` to make no request; the revocation status is then `not-checked`.

## Word, Excel and Visio

By default Word, Excel and Visio open, render, edit and save documents locally.

- linked (external) pictures are not fetched (Word and Excel download them on open, so this is
  stricter than those applications; see
  [docs/powerpoint-external-content.md](docs/powerpoint-external-content.md));
- external workbook links, Visio data record sets and linked chart data are kept but never
  refreshed; Excel web queries and `WEBSERVICE` are not evaluated;
- fonts come from the device and the file only, unless the host sets `googleFonts` (below);
- workers and application code are loaded from the page's own origin. Opening Word's Design tab
  loads its gallery code and built-in presets from that origin, with no document data in the request.
  Opening or inserting a TIFF
  picture in Visio can load its bundled decoder on demand. The request carries no picture bytes
  or filename. If the decoder is unavailable offline, the drawing still opens and the TIFF is
  omitted from the preview. Rendering or editing Office Math can load bundled equation code from
  the same origin. These requests carry only the application asset URL. The file still opens and
  retains its equations if that code is unavailable.
- Loading a Visio drawing with an embedded workbook chart or an admitted metafile can request
  the corresponding bundled reader from the app's own origin. Excel loads its bundled XPath
  code when a workbook contains FILTERXML or the formula is entered, including in a conditional
  formatting rule. These requests send the asset URL and standard browser metadata. Drawings and
  workbooks without those features skip the requests. If the assets are unavailable offline, the
  drawing reports an unsupported preview or the formula returns an error. File contents stay local.
- Selecting a local stencil file or running a Visio stencil authoring command loads its bundled
  translations, editing code and worker from the application's own origin as needed. Those
  requests carry only the application asset URLs. The Shapes pane menus are in the app bundle.
  Stencil files and My Shapes stay in session memory. Save Stencil writes a local download.
  Hosts can omit the Shapes pane, and leaving the stencil commands unused makes no stencil
  asset requests. If an asset is unavailable offline, the command reports the failure.
- Word's Mailings commands load editor code on demand from the app's origin. Choosing a local
  recipient file also loads a same-origin parser worker where supported. These requests carry no
  file name, recipient data or document bytes. Hosts can disable Mailings with the editor's
  read-only setting or by disabling editing in its customization. Use the app's cached files
  offline. If workers are unavailable or blocked by the host's `worker-src` policy, the editor
  reads the file locally on the page. Recipient lists and selections stay in memory. The editor
  calls a host's `mailMergeDataSource.read()` only after **Use host recipients**, and the host controls any
  requests it makes. In the Office suite, **Edit Individual Documents** creates a new local
  profile document containing the merged text. It uses the suite's existing document storage.
- Excel requests the selected non-English language file from the app's own origin when the
  user or host selects that display language. English is built in and is used until the file
  loads, or whenever it is unavailable. Excel also loads bundled chart and PivotTable dialog
  code from that origin when the associated controls are used. These requests send only the
  asset URL and standard browser request metadata. Keep `locale` at `en` to skip language
  requests, and leave a dialog or pane unopened to skip its request.

**Fonts, when the host or the reader turns them on.** `<docx-editor>`, `<xlsx-editor>` and
`<visio-viewer>` take the same `googleFonts` value as PowerPoint, but unset leaves it to the
reader's File > Options > Trust Center toggle "Download fonts that are not installed or embedded",
which is off by default (a host value locks the toggle). With
`{ baseUrl: 'https://fonts.googleapis.com/css2' }` (or a self-hosted mirror) they send the same
requests PowerPoint does, from the same shared code (`ooxml-core/text`, `ooxml-ui/fonts`): the
public catalogue families (or metric-compatible clones such as Carlito for Calibri) the file
declares (theme, styles, font table) and the device lacks, then the font files for the faces it
draws. When downloads are enabled, choosing a new theme or font scheme in Word can cause the same
requests during editing. The file, its text and the names of fonts outside Google's public
catalogue are never sent. `false` or an invalid `baseUrl` sends nothing.

Hyperlinks and in-product help links open only when the reader follows them.

## Teams (`ooxml-core/teams`, `ooxml-ui/teams`, the `openteams-*` packages)

- **Sync and signaling:** chat, presence and session data go only to the `syncUrl` and
  `signalingUrl` the host or the user configures. Local mode stays in the browser
  (`BroadcastChannel`). The reference server in `viewers/teams/server` makes no outbound calls.
- **Calls:** without an `iceServers` setting, calls use Google's public STUN server
  `stun:stun.l.google.com:19302`, which learns each participant's IP address. Set `iceServers`
  to your own STUN/TURN servers, or to `[]` to use none. Audio and video flow between peers or
  through the TURN server you configure.
- **Call signaling:** offers, answers and network candidates (which carry each participant's IP
  addresses) go through the `signalingUrl` server to the other people in the call. With accounts the
  reference server relays them only between signed-in users speaking as themselves, and a private
  chat's call only between that chat's members; a server with only `authorizeSync` relays none.
- **Website tabs:** a website tab loads the URL a member added, in a frame, for everyone who
  opens it. A host's Content Security Policy decides which sites may be framed: the OOXML Office
  suite frames none, so there the tab offers "Open externally", which opens the site in a new
  browser tab.

## OOXML Office suite (`apps/office-suite`, the GitHub Pages site and the desktop app)

The suite uses the editors above, so their sections apply. In addition:

- **One script policy for the whole site:** the GitHub Pages origin is shared with other project
  sites, so every published page (the suite, its 404 page, every viewer's documentation and demos)
  and the desktop app carry the same Content Security Policy from `scripts/csp.mjs`: script only
  from the site itself, never inline, and frames only from the site and the online-video players
  (`www.youtube.com`, `player.vimeo.com`). The Pages build fails on a page without it.
- **Fonts:** Settings > Fonts chooses Google Fonts (the default), your own font server, or no
  downloads, for PowerPoint, Word, Excel and Visio alike. The policy allows only Google's two
  font hosts; a deployment that uses its own font server adds that server to it.
- **AI assistant:** off the network by default (the shipped `ai-config.json` selects the offline
  assistant). When the user or the deployment's policy selects a provider, the prompt, the
  conversation and the document context it needs are sent to that provider, under the
  provider's own terms. The context is the files the user attached; hidden and very hidden
  sheets and the cells of hidden rows and columns are left out unless the deployment's policy sets
  `context.includeHidden`. Keys stay in this tab's session storage only and are cleared when it
  closes; keys an earlier version remembered in localStorage are deleted when the app loads. A
  saved model is used only while the policy still allows its provider. See
  [docs/suite-ai-config.md](docs/suite-ai-config.md).
- **OneDrive:** off unless the host configures it. Files and the access token then go to
  Microsoft Graph. See [docs/onedrive-integration.md](docs/onedrive-integration.md).
- **Shared AI chats:** sync over the configured collaboration or Teams server, otherwise stay in
  this browser. The room, and so the sync server, receives each request's text, each reply's
  text as shown (which can quote attached files), the model's label, the names and kinds of
  attached files and the number of edits and new files a reply proposed. The files themselves,
  the reply's raw text, the proposed edits with their before-and-after text, and file ids stay on
  the asker's device. When the deployment's `ai-config.json` sets `collaboration.credentialsUrl`
  (off by default), the suite sends a `GET` with the user's cookies to that URL right before a
  shared chat connects, and again before each reconnect or once the token expires, to get a
  short-lived token for the sync server. No document or chat content is sent, and the token is
  kept in memory only. Leave the field empty to turn the request off; a sync server that needs a
  token then cannot be used.
- **Updates:** the installed app checks its own origin for a new `sw.js` at start, on focus and
  hourly. The service worker caches only the app's own files, and serves a saved copy only while
  it matches the SHA-256 the build recorded for it; a copy that does not is deleted and fetched
  again from the site. A new version installs only when every file it downloads matches, and
  otherwise tries again at the next check. After an update, the previous version's files stay
  cached for windows still running it and are deleted once none is. Each version's cache also
  keeps its file list and an anonymous marker per window it served (a hash of the browser's
  window id), so a later version can tell which saved files those windows may use. The shell's
  script and style tags carry subresource integrity. Excel's non-English language files are
  saved only after they are requested and match their recorded hashes. Languages already saved
  for that version can be used offline.
- **Desktop:** the Tauri shell has no updater, telemetry or plugins. It reads and writes the
  files the user opens, and lists the file names inside a folder the user connects so they can be
  found; nothing is sent. Its window only shows the bundled app: a web or `mailto:` link the user
  follows opens in the system's default browser or mail client, and other destinations (another
  site in the window, `file:` addresses, new in-app windows) are refused. The operating system's
  WebView (WebView2, WKWebView) may send its vendor's own diagnostics under the operating
  system's settings.

## MCP servers (`mcp/` and the viewers' MCP packages)

The MCP servers make no network request of their own: they read and write files under the root
they are given and answer over stdio. On Windows, a file path naming a network share can make the
system contact that host while the path is checked, before it is refused. The Excel tools refuse
formulas that would make Excel or another spreadsheet application fetch data, open links or call
an AI service when the workbook is opened (see [docs/mcp.md](docs/mcp.md)).

## Data kept on the device

Nothing below is sent anywhere by this code. Clearing the site's data removes it.

- **PowerPoint:** autosave and recovery snapshots, the session, presenter and audience state in
  IndexedDB (`pptx-viewer-autosave`, `pptx-viewer-session`, `pptx-viewer-presentation`,
  `pptx-viewer-audience`); preferences, including Trust Center choices, in localStorage
  (`pptx-viewer-prefs`). File > Account > Storage & Privacy shows and clears them. Fonts the
  user adds under File > Options > Custom fonts are kept for the session only.
- **Word, Excel, Visio:** small preferences in localStorage (view options, custom lists, saved
  shapes, Excel's Error Checking rules in `ooxml.xlsx.errorCheckingRules`), including the Trust
  Center font-download choice (`ooxml.docx.allowFontDownloads`, `ooxml.xlsx.allowFontDownloads`,
  and `allowFontDownloads` in Visio's `ooxml-visio-options`).
- **Teams:** chat snapshots, read state, drafts, presence and follows in the storage the host
  provides; the server configuration in localStorage. Session tokens are kept in memory only.
- **Word document design:** Set as Default writes the document theme palette and names, font
  names, a built-in effect choice and paragraph spacing to the host's preference store.
  Embedders enable this by assigning `wordDesignPreferences` and disable it by clearing that
  property. The suite stores these values on this device by profile under
  `ooxml-suite-word-design-<profile id>` in localStorage. Saving requires the explicit command,
  and Word reads the values when creating a blank document. In suite Settings > General, turn
  off Remember my default Word design to stop reading and writing this preference. Nothing is sent.
- **Word drawing layout:** Wrap Text > Set as Default Layout stores one wrap choice per drawing
  kind under `ooxml-word-drawing-wrap.picture` and `ooxml-word-drawing-wrap.shape` in localStorage.
  Word reads the choice when inserting pictures or drawing shapes. The device stores only the
  explicitly selected wrap choice. Clear these keys to reset the defaults. If the host blocks
  localStorage, the choice lasts until the page closes.
  Arrange commands may load bundled editor code from the app's origin when invoked. These asset
  requests carry no drawing or document data.
- **Suite:** the document library, attachments, versions and chat history in IndexedDB per
  profile; profiles, favourites, theme (`ooxml-suite-theme`), fonts (`ooxml-suite-fonts`), the
  other Settings (`ooxml-suite-settings`) and AI preferences in localStorage; AI keys in session
  storage only. Settings > Privacy shows how
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

### Offline PowerPoint symbol fallback

When Wingdings or Symbol is unavailable, the PowerPoint renderer translates known
characters to Unicode and tries the paragraph font, then a bundled Noto Sans
Symbols 2 subset (U+2000-2BFF). The subset is inline WOFF2 data, enabled without
font downloads. It makes no outbound request, sends no document or font names,
and creates no persistent storage. Image/print exports collect the same font
rule. Host CSP must permit `font-src data:` and the viewer's inline styles; the
published pages' and desktop policy (`scripts/csp.mjs`) already does. The font-download setting does not disable
this local fallback. This is approximate symbol substitution, not exact Wingdings
outlines; unmapped symbol bullets use a plain bullet.

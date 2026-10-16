# OOXML Office application

The suite mounts the real Word, Excel, PowerPoint, Visio and OpenTeams components in one application. `apps/office-suite/src` owns host UI and framework wiring; `src/ui/src/suite` owns browser document persistence. Office parsing, editing and serialization remain in `src/core`.

## Account and workspace UI

The app header contains a native popover account manager, app launcher, search and shared assistant. The account manager edits the local display name, optional contact email and job title, photo, and workspace label. Teams receives the same stable user id and display name. Email and photo remain local. Additional local profiles use separate document databases, favourites and Teams workspace ids. Switching flushes and saves open editors before reloading; the default profile preserves the previous library and Teams identity.

Local profiles are organizational conveniences, not authentication or authorization boundaries. No Microsoft sign-in, organization membership, verified email or SharePoint permissions are implied. The hosted preview's access control is separate from the app's local profiles. Appearance currently remains a browser-wide preference.

`node scripts/suite-profile-check.mjs` checks editing/persistence, photo controls, Teams identity, saved changes across profile switches, separated libraries, favourites, the launcher, mobile search and responsive account menus.

## Implemented workflows

- Import native DOCX, XLSX, PPTX and VSDX files into one IndexedDB library. One persistent document id maps to one live editor tab. Tabs preserve edits while switching applications. Save and Download serialize the current editor model.
- Create blank documents, workbooks and presentations using the existing engines.
- Share a local document reference in Teams, reopen it in its existing editor and return to the captured channel. Local channel links point to the same saved document. Download actions are intercepted by the host so they return file bytes, not the launcher HTML. Pinned file tabs use the same opening contract.
- With a configured Teams server, upload through its authenticated storage endpoint and save edited documents as new channel copies. Existing remote attachments are imported as local copies; they are never silently overwritten.
- Open native Office packages from a parent's `embeddings` directory. Saving updates child and parent atomically, with revision checks. Reopen a child after its parent changes. Binary OLE containers are offered for download only. Updating a chart workbook does not regenerate every external chart cache or object preview.
- Open on **Chat**, the suite's front door: a conversation history (search, pin, rename, delete, grouped by date, kept per profile in IndexedDB), a model picker and starters for each app. Chat and the assistant pane beside an open document are the same chat, so a conversation moves between them ("Open in Chat"). The pane attaches the open document and suggests prompts for its app; **Ask AI** in the document command bar opens it.
- Attach files explicitly (and optionally the current Teams conversation). The assistant reads current saved editor content and replies in Markdown; anything it wants done travels in one fenced `ooxml-actions` block that becomes cards. Edit cards propose bounded Word run, spreadsheet cell, PowerPoint run and Visio shape-text edits; applying one updates all targeted documents atomically, and Undo restores the prior snapshots if no intervening edit occurred. Create cards build a new Word document (title, headings, paragraphs, bullet and numbered lists), workbook (sheets of values and formulas) or deck (title slide and bulleted slides with notes) through `ooxml-core` and open it in a tab. AI proposals for embedded documents are read-only; use the editor for parent save-back.
- Light, Dark and System appearance plus four accent themes. Preferences persist; the host updates open editors and Teams. Document page colours remain document content.

## Build and verification

`bun run build:suite` builds core, shared UI and a self-contained `suite-dist` application. `bun run preview:suite` bundles and serves the launcher (build core and UI first). `scripts/build-pages.mjs` includes the same suite bundle alongside the viewer documentation. Desktop builds consume `suite-dist` through Tauri.

Run `bun run test:suite` against the preview URL (default localhost:8128, override with `SUITE_URL`). The browser check uses actual editors and round-trips downloaded files. Only the external AI provider is stubbed; it asserts the supplied Word and Excel context, applies a proposal and undoes it. Unit tests cover package preservation, missing embedded targets, atomic storage conflicts, presentation text edits and Teams host events.

## Boundaries

Files are local to the current browser/WebView profile. This is not a Microsoft 365 account integration or cloud drive. Multi-device chat/calls require a configured Teams sync/signaling server and appropriate TURN configuration. Office document coauthoring is not wired to suite file identities yet. The suite has conflict detection, not live concurrent document merging.

The assistant is governed by deployment configuration: `ai-config.json` beside the app (or `globalThis.ooxmlAiConfig` from a host) chooses the providers and models (Anthropic, OpenAI, GitHub Models, Azure OpenAI, Gemini, Mistral, Groq, DeepSeek, xAI, Together, Fireworks, Perplexity, OpenRouter, Ollama, LM Studio, an organisation gateway or any OpenAI-compatible endpoint), whether users may add their own, which features are on, the context and history limits and the instructions every request carries; the UI only offers what it allows. See [configuring the assistant](suite-ai-config.md). Replies stream and can be stopped; regenerating keeps the earlier reply as a version. Users reference open tabs and library files with `@` or Attach. Proposed edits show a before-and-after diff and can be applied one by one; every save and assistant edit keeps the replaced bytes as a version (Versions in the command bar, or the file menu), which can be restored or downloaded. Keys a user enters stay in the tab's session storage unless they choose to remember them. Until a model is connected, an **offline preview** answers from simple rules on the device; it says so in every reply, describes attached files from their real content, drafts sample files and can title-case an attached document or deck, so the whole flow can be tried without a key. Context is bounded (Word: 500 paragraphs; Excel: A1:AD100 on the first four sheets; Teams: 50 current-channel messages; the policy's context limit). The application never executes model-generated code or silently posts AI output to channels; nothing is written until the user presses a card's button.

Unsupported Office features remain governed by each editor's compatibility diagnostics. No Office parity or universally lossless export is claimed. Native file associations, OS save dialogs, signing, notarization and updates remain desktop release work. Windows and macOS use the same application bundle; macOS packaging must run on a Mac.

## Visual references

The interface follows Microsoft's document-oriented navigation, restrained surfaces, compact commands and persistent side panes, with original icons and OOXML branding:

- [Microsoft 365 start-page reference](https://learn.microsoft.com/en-us/microsoft-365/document-collaboration-partner-program/)
- [Editing files within Teams](https://support.microsoft.com/en-gb/teams/files/edit-a-file-in-microsoft-teams)
- [Fluent typography](https://fluent2.microsoft.design/typography/)
- [Fluent layout](https://fluent2.microsoft.design/layout)

OOXML Office is independent of Microsoft.

## Application package and installation

`apps/office-suite` is the private workspace package `ooxml-office`. This is the application source used by the web and Tauri builds, not a demo wrapper or another Office logic library. It is not published to npm. Standalone app entry points are generated from the same source; each has a distinct manifest id and install scope under `apps/word/`, `apps/excel/`, `apps/powerpoint/`, `apps/visio/` and `apps/teams/`. They use one profile identity, library and theme preference on the same origin. Their app launcher links back to Office or the other standalone apps.

The service worker precaches generated shell/editor resources for offline use. It does not cache arbitrary network traffic, account endpoints, assistant calls or Teams server data. Installation requires a secure origin, and offline availability starts after initial asset caching succeeds. Updates wait for existing app windows to close. Offline chat remains local; this does not make network services available offline. The account menu provides installation guidance for browsers, including Safari on iOS. Desktop WebViews retain their bundled assets without depending on PWA installation.

General Teams attachments use a separate profile-scoped binary store. PDF, image, archive and extensionless files can be sent and downloaded, up to the existing 32 MB limit per file. Native Office files still open in the integrated editors. Passive image/audio/video formats have an inline preview; other formats download. Original filenames are preserved separately from unique server storage keys.

PowerPoint's plain DOM and body-mounted dialogs inherit one suite palette. Global shell controls explicitly exclude PowerPoint descendants. The Options theme choice inherits the Office theme; appearance is controlled from the suite header. Presentation creation no longer imports the unrelated combined automation entry first, and editor loading starts alongside document creation or on pointer/focus intent. Live document tabs are retained when switching apps.

`node scripts/suite-regression-check.mjs` verifies attachment bytes/names after reload, composer focus and emoji insertion, phone overflow, stable navigation width, light/dark PowerPoint settings, dialog cleanup, properties field bounds, distinct manifests and offline app reopening.

## Workspace, storage and tab actions

The account avatar opens profile and appearance controls. The workspace switcher selects a local workspace; the sidebar ellipsis manages storage and connected folders. Phone and standalone layouts expose the same storage controls above the file list.

Search waits briefly after typing and searches connected folders asynchronously. Aborted/stale requests cannot replace newer results. Searches match filenames and relative paths, not file contents. Traversal is capped at 5,000 files, 20,000 entries and 40 levels per connected folder, with truncation/access issues shown in the UI. At most 100 disk matches render at a time. There is no unrestricted disk crawler or OS-wide index.

Browsers supporting the File System Access API retain chosen directory handles per local profile, subject to browser permission renewal. Others offer a selected-folder/file snapshot that lasts for the current page session. The desktop app uses a native folder picker and background Rust scan, and can reveal known sources in Explorer or Finder. Browsers do not expose an absolute disk path or a native reveal action. Imports through the ordinary upload input cannot recover their original source path.

Disk results import an editable workspace copy; saving does not overwrite the source file. File actions expose its location, download, and removal from the workspace. Removal deletes the app copy and cached embedded children transactionally, with immediate Undo. Original files stay on disk. Tabs support middle-click and context actions to close, close others, close to the right and close all documents. Edited tabs save before closing; failures leave the tabs open. The Teams tab can also close while retaining its conversation/draft state.

`scripts/suite-files-check.mjs` checks the distinct menus, connected-folder enumeration, stored folder identity, stale search results, source preservation, remove/undo and tab edits. It uses generated OPFS handles for the folder picker: native OS permission prompts and persisted real-disk handle renewal require platform testing. The test avoids deserializing OPFS fixture handles due to a reproducible Chromium crash unrelated to application code. Rust tests cover authorized-root checks, path traversal rejection and file scanning.

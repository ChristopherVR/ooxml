# OOXML Office desktop

**The OOXML Office suite as a desktop app: a private Tauri 2 shell around the same bundle and the same five editors the web build ships.**

This package (`ooxml-desktop`) is private and not published. It adds a few narrowly scoped native commands to the [Office suite](../apps/office-suite/README.md): choosing a folder, reading and searching inside that folder, and revealing a chosen file in Explorer or Finder. There are no disk write or delete commands.

## Build

1. Install the Tauri prerequisites: Rust, plus WebView2 and the Windows compiler on Windows, or Xcode on macOS.
2. From the repository root: `bun install`.
3. From `desktop/`: `bun install`, then `bun run build`. The before-build step builds the core, the UI and the suite assets.

For development, build the suite once (`bun run build:suite` at the root), then run `bun run dev` here. `bun run check` runs `cargo check` on the Rust side.

Windows produces an NSIS installer; macOS produces an app and a DMG when built on a Mac. The output is under `src-tauri/target/release/bundle/`.

## Security model

- **Folders by consent.** A native folder picker must authorise each root in the running app; arbitrary paths, traversal and resolved links outside it are refused. Authorisation is renewed through the picker after a restart.
- **The window shows only the bundled app** (`src-tauri/src/navigation.rs`). A web or `mailto:` link opens in the system browser or mail client; navigating the window anywhere else, `file:` and `ftp:` addresses and new in-app windows (`window.open`) are refused.
- **One Content Security Policy.** The app carries the same policy as the published site (`scripts/csp.mjs`), kept equal to `tauri.conf.json` by `scripts/csp.test.mjs`.

## Privacy and network use

The shell has no updater, telemetry or plugins. It reads the files the user opens and lists the file names inside a connected folder so they can be found; nothing is sent. AI and remote Teams connections still need a network and a provider or server you configure. The operating system's WebView (WebView2, WKWebView) may send its vendor's own diagnostics under the system's settings. See [PRIVACY.md](../PRIVACY.md#ooxml-office-suite-appsoffice-suite-the-github-pages-site-and-the-desktop-app).

## Status and limits

- These are unsigned development builds: no code signing, notarisation, automatic updates or file associations.
- Files live in the WebView's IndexedDB, separate from your browser profile. Import and download are explicit; the app does not overwrite arbitrary disk files or synchronise a cloud drive.
- Folder scans run on a background thread with bounded traversal. Search is by file name and relative path, not document contents or a system-wide index.

See [suite integration](../docs/suite-integration.md) for the supported workflows and remaining limits.

## License

[Apache-2.0](../LICENSE).

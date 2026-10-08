# OOXML Office desktop

The private Tauri 2 host bundles the same suite and all five editors used on the web. It exposes narrowly scoped commands for selecting a folder, reading/searching within that folder and revealing a selected file in Explorer or Finder. A native picker must authorize each root in the running app; arbitrary paths, traversal and resolved links outside it are rejected. No disk write/delete commands are exposed.

Install the Tauri prerequisites (Rust, WebView2 and the Windows compiler, or Xcode on macOS). From the repository root run `bun install`, then from `desktop/` run `bun install` and `bun run build`. The before-build command builds the core, UI and suite assets. For development, build the suite once before `bun run dev`.

Windows produces an NSIS installer; macOS produces an app and DMG when built on a Mac. Artifacts are under `src-tauri/target/release/bundle/`. These are unsigned development builds; no signing, notarization, automatic updater or file associations are configured.

The editor code is bundled locally. AI and remote Teams connections still require a network and a configured provider/server. Files use the WebView's IndexedDB and are separate from your browser profile. Import and download are explicit; the app does not overwrite arbitrary disk files or synchronize a cloud drive.

See [suite integration](../docs/suite-integration.md) for supported workflows and remaining limits.

Folder scans run on a background thread, with bounded traversal. Native folder authorization is renewed with the folder picker after restarting the app. Search is by filename and relative path, not document contents or a system-wide index.

# @christophervr/pptx-viewer

Two things under one name: an `npx` installer that sets up the PowerPoint viewer for your framework, and, as a library, the React viewer re-exported from [`pptx-react-viewer`](https://www.npmjs.com/package/pptx-react-viewer).

[![npm version](https://img.shields.io/npm/v/%40christophervr%2Fpptx-viewer.svg)](https://www.npmjs.com/package/@christophervr/pptx-viewer)
[![license](https://img.shields.io/npm/l/%40christophervr%2Fpptx-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)

[Live demo](https://christophervr.github.io/ooxml/pptx/demo/) · [Docs](https://christophervr.github.io/ooxml/pptx/guide/installation) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/cli/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![The interactive installer selecting React and MCP, then scaffolding a starter app](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/pptx/.github/assets/packages/cli-installer.webp)

## The installer

```bash
npx @christophervr/pptx-viewer@latest
# or
bunx @christophervr/pptx-viewer@latest
```

It asks what you are building, checks your project, then installs the right package or scaffolds a new starter app. Nothing needs installing first.

| Target     | Installs                                                                                                    |
| ---------- | ----------------------------------------------------------------------------------------------------------- |
| React      | [`pptx-react-viewer`](https://www.npmjs.com/package/pptx-react-viewer), for React 18.2 or 19                |
| Vue        | [`pptx-vue-viewer`](https://www.npmjs.com/package/pptx-vue-viewer), for Vue 3.5 or newer                    |
| Angular    | [`pptx-angular-viewer`](https://www.npmjs.com/package/pptx-angular-viewer), for Angular 19 to 22            |
| Svelte     | [`pptx-svelte-viewer`](https://www.npmjs.com/package/pptx-svelte-viewer), for Svelte 5                      |
| Vanilla JS | [`pptx-vanilla-viewer`](https://www.npmjs.com/package/pptx-vanilla-viewer), no framework                    |
| Core       | [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core), the engine with no UI                 |
| MCP        | Nothing: it prints the client config for [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp) |

How it works:

1. **Pick targets.** Arrow keys move, Space toggles, `a` selects all, Enter confirms. Choose at most one UI framework, plus `core` and `mcp` if you like. Without raw keyboard input it falls back to a numbered prompt.
2. **Compatibility check.** With a `package.json` present, it compares your installed `react`, `vue`, `@angular/core` or `svelte` major with what the viewer supports, and asks before going on when they differ.
3. **Install here, or scaffold.** Installing here needs an existing `package.json`. Scaffolding creates a project with `create-vite` (React, Vue, Svelte, Vanilla JS) or `@angular/cli` (Angular, whose CLI needs Node.js 22.22.0+, 24.13.1+ or 26.0.0+), replaces its entry file with a working viewer, installs the viewer, the engine and the framework's companion packages, offers collaboration (`yjs`, `y-websocket`, `y-webrtc`) and, for Vanilla JS, 3D (`three`), then starts the dev server.

It detects `bun`, `pnpm`, `yarn` or `npm` from your lockfile, then from the manager that launched it, and prints a quick-start snippet when it is done.

### Non-interactive use

```bash
npx @christophervr/pptx-viewer --target react,mcp --yes
npx @christophervr/pptx-viewer --target vue --scaffold --dir my-app --yes
npx @christophervr/pptx-viewer --pm pnpm
```

| Flag             | Meaning                                                                                          |
| ---------------- | ------------------------------------------------------------------------------------------------ |
| `--target <ids>` | Comma-separated `react`, `vue`, `angular`, `svelte`, `vanilla`, `core`, `mcp`; skips the picker. |
| `--scaffold`     | Create a new project instead of installing here; needs exactly one UI target.                    |
| `--dir <name>`   | The new project's folder (default `pptx-<target>-app`).                                          |
| `--pm <manager>` | `bun`, `pnpm`, `yarn` or `npm`, instead of detecting one.                                        |
| `--yes`, `-y`    | Skip confirmations, including the compatibility warning; optional prompts take their default.    |
| `--help`, `-h`   | Print usage and exit.                                                                            |

Without `--target` and without a terminal (piped stdin, most CI runners) it exits with an error instead of waiting for input. An unknown flag, target or manager exits with a one-line error, and Ctrl+C exits with status 130. Colour turns off under `NO_COLOR` or when output is not a terminal.

## As a library

```bash
npm install @christophervr/pptx-viewer pptx-react-viewer
npm install react react-dom framer-motion lucide-react react-icons jspdf jszip fast-xml-parser i18next react-i18next
```

```tsx
import { PowerPointViewer } from '@christophervr/pptx-viewer';
import 'pptx-react-viewer/styles';

// <div style={{ height: '100vh' }}><PowerPointViewer content={bytes} canEdit /></div>
```

Everything [`pptx-react-viewer`](https://www.npmjs.com/package/pptx-react-viewer) exports from its root is re-exported here, so its [README](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/react/README.md) and [docs](https://christophervr.github.io/ooxml/pptx/react/) apply as written. The stylesheet and the subpath entries (`/ai`, `/i18n`, `/viewer`) ship under the real package name, which is why the install line adds it. For a new React project, depending on `pptx-react-viewer` alone is simpler; the peers here are marked optional so that using only the installer never warns about React.

## Supported environments

- Node.js 22.3 or newer to run the installer and to build.
- The library inherits the React viewer's support: React 18.2 or 19, and the last two major versions of Chrome, Edge, Firefox and Safari ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).

## Privacy and network use

The installer runs your package manager and, when scaffolding, `create-vite` or `@angular/cli`, which download packages from your configured registry. It sends nothing else.

As a library this package is the React viewer, which renders and saves decks in the browser and never uploads them. It contacts other servers only to download web fonts that are neither installed nor embedded (Google Fonts by default; `googleFonts: false` or a mirror turns it off), to fetch a deck's linked pictures and media after the reader chooses Enable Content (or `allowExternalContent: true`), to open a hyperlink the reader clicks, and for Share and Broadcast once started (your server, or y-webrtc's public signaling and STUN servers peer to peer). The [React viewer's README](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/react/README.md#privacy-and-network-use) gives the details. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

The installer's targets above, plus the [project overview](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) for how the PowerPoint packages fit together.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.

# openteams-solid-viewer

**OpenTeams for SolidJS: a team workspace (channels, chat, presence and meetings) in one component, connected to a server you run.**

[![npm](https://img.shields.io/npm/v/openteams-solid-viewer.svg)](https://www.npmjs.com/package/openteams-solid-viewer)
[![license](https://img.shields.io/npm/l/openteams-solid-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE)
[![types](https://img.shields.io/npm/types/openteams-solid-viewer.svg)](https://www.npmjs.com/package/openteams-solid-viewer)
[![status: beta](https://img.shields.io/badge/status-beta-orange.svg)](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)

[Live demo](https://christophervr.github.io/ooxml/teams/demo-solid/) · [Documentation](https://christophervr.github.io/ooxml/teams/) · [Solid guide](https://christophervr.github.io/ooxml/teams/frameworks/solid) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/solid/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

![The OpenTeams workspace: a channel list, a conversation with reactions and a quoted reply, and the composer](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/teams/docs/public/hero-workspace.webp)

OpenTeams is not Microsoft Teams and is not affiliated with Microsoft. There is no end-to-end
encryption: your server can read chat and files.

## Install

```bash
npm install openteams-solid-viewer
```

Peer dependency: `solid-js` `^1.9.0` (Solid 2 is not supported yet).
[`ooxml-core`](https://www.npmjs.com/package/ooxml-core) (the workspace logic),
[`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui) (the `<teams-app>` element) and `lit` come
with it. Use one OpenTeams binding per page: they register the same custom elements.

## Quick start

```tsx
import { render } from 'solid-js/web';
import { Teams } from 'openteams-solid-viewer';

const config = {
	mode: 'server' as const,
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
};

render(
	() => (
		<div style={{ height: '100vh' }}>
			{Teams({
				workspaceId: 'acme',
				userName: 'Ada',
				config,
				onOpenFile: (detail) => console.log(detail.attachment.kind, detail.url),
			})}
		</div>
	),
	document.getElementById('app')!,
);
```

`Teams(props)` is a plain function that returns the `<teams-app>` element, so it needs no JSX
transform: call it inside JSX or pass it to `render`. Props are applied reactively (use getters for
values that change) and handlers are read when the event fires. During server rendering it emits
the empty element, which the browser claims while hydrating. Give the container a height.

`config` points at a server that speaks the OpenTeams contracts, such as
[`openteams-server`](https://www.npmjs.com/package/openteams-server) (locally:
`npx openteams-server --insecure-dev`, then `ws://127.0.0.1:8787/sync` and
`ws://127.0.0.1:8787/signal`). When the server requires accounts, the workspace asks the person to
sign in. Leave `config` out and the app uses the settings saved in its dialog, or runs in local
mode, where tabs of one browser share the workspace over `BroadcastChannel`.

## Props

| Prop          | Type                                                  | Notes                                                                                                                      |
| ------------- | ----------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| `workspaceId` | `string`                                              | The shared room: 1 to 128 letters, digits, `-` or `_`. Default `demo`.                                                     |
| `userName`    | `string`                                              | Display name.                                                                                                              |
| `userId`      | `string`                                              | Stable id for authorship; by default one remembered in this browser.                                                       |
| `config`      | `TeamsServerConfig \| null`                           | `mode`, `syncUrl`, `signalingUrl`, `iceServers`; optional `iceTransportPolicy`, `token`, `sessionToken`. `null` clears it. |
| `locale`      | `string`                                              | BCP 47 tag such as `de` or `zh-CN`; empty or unknown means English.                                                        |
| `uploadFile`  | `(file, { workspaceId, signal }) => Promise<{ url }>` | Where attachments go; by default the server's `/files` endpoint.                                                           |
| `openers`     | `{ docx?, xlsx?, pptx?, vsdx?, other? }`              | One function per file kind, called with the open-file detail.                                                              |
| `embeds`      | `{ docx?, xlsx?, pptx?, vsdx?, other? }`              | One function per file kind returning a viewer page URL for the preview pane.                                               |
| `class`       | `string`                                              | Classes for the `<teams-app>` element (`className` works too).                                                             |

The returned element has a read-only `client` (the core client) and `previewContent(detail)`.
Full reference: [element API](https://christophervr.github.io/ooxml/teams/api).

## Events

| Prop             | Called with       | When                                                                             |
| ---------------- | ----------------- | -------------------------------------------------------------------------------- |
| `onOpenFile`     | `(detail, event)` | Someone opened an attachment; call `event.preventDefault()` to open it yourself. |
| `onReady`        | `({ user })`      | The workspace is up and knows who the user is.                                   |
| `onConfigChange` | `({ config })`    | The person saved new server settings in the app.                                 |

## Your own UI

`createTeamsClient` gives you the same client without the bundled UI, as signals.

```tsx
import { For } from 'solid-js';
import { createTeamsClient } from 'openteams-solid-viewer';

function Channels() {
	const { client, state } = createTeamsClient(() => ({
		workspaceId: 'acme',
		user: { id: 'ada', name: 'Ada' },
		config,
	}));
	return (
		<For each={state()?.channels ?? []}>
			{(c) => <button onClick={() => client()?.select(c.id)}>{c.name}</button>}
		</For>
	);
}
```

The client is recreated when the options change and destroyed with its owner; return `null` from
the options function to hold off connecting. The actions are listed under
[the raw client](https://christophervr.github.io/ooxml/teams/api#the-raw-client).

## Theming

The element styles itself in its shadow root, so there is no stylesheet to import. Colours,
spacing and radii come from the shared `--office-*` tokens; the OpenTeams ones are prefixed
`--office-teams-` (such as `--office-teams-brand`, `--office-teams-rail` and
`--office-teams-panel`). Override them on `teams-app` or an ancestor. The app follows the system
colour scheme or `data-office-theme="light"` / `"dark"` on the page, and each person can pick
System, Light or Dark in its settings. See
[theming](https://christophervr.github.io/ooxml/teams/theming).

## Localisation

The UI ships in English, German, Spanish, French and Simplified Chinese: set `locale` to `de`,
`es`, `fr` or `zh-CN`; an empty or unknown tag means English. Text people wrote is never
translated.

## Status and limits

OpenTeams is **Beta** ([stability](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)).

- No end-to-end encryption: the server stores and can read every message and file.
- Calls are full mesh with no media server, comfortable up to about a dozen people; strict
  networks need a TURN server you run.
- No SSO, password reset, recording, background blur, live captions, or push and email
  notifications (notifications work only while the app is open).
- Markdown is a safe subset, not full CommonMark or GFM.

See [limitations](https://christophervr.github.io/ooxml/teams/limitations),
[teams-parity.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-parity.md) and
[parity-checklist.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Privacy and network use

- **Sync and signaling** go only to the `syncUrl` and `signalingUrl` that you or the person
  configure; local mode stays in the browser. Call signaling carries each participant's network
  addresses to the others in the call.
- **Calls** use Google's public STUN server (`stun:stun.l.google.com:19302`) when no `iceServers`
  are set, which learns each participant's IP address. Set `iceServers` to your own STUN or TURN
  servers, or to `[]` for none.
- **Website tabs** load the URL a member added, for everyone who opens the tab.
- **On the device:** chat snapshots, read state, drafts and the server settings in localStorage;
  session tokens stay in memory. Nothing collects analytics or telemetry.

See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#teams-ooxml-coreteams-ooxml-uiteams-the-openteams--packages).

## Browser support

The last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least
([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
Cameras and microphones need `localhost` or HTTPS. SolidJS 1.9; Node.js 22.3 or newer for builds
and tests.

## Documentation and demos

- [Documentation](https://christophervr.github.io/ooxml/teams/),
  [Solid guide](https://christophervr.github.io/ooxml/teams/frameworks/solid) and
  [Solid demo](https://christophervr.github.io/ooxml/teams/demo-solid/).
- Server: [`openteams-server`](https://www.npmjs.com/package/openteams-server) and
  [bring your own server](https://christophervr.github.io/ooxml/teams/server). Use the same major
  version as this package: both speak the same Teams protocol.
- Other frameworks: [React](https://www.npmjs.com/package/openteams-react-viewer) ·
  [Vue](https://www.npmjs.com/package/openteams-vue-viewer) ·
  [Angular](https://www.npmjs.com/package/openteams-angular-viewer) ·
  [Svelte](https://www.npmjs.com/package/openteams-svelte-viewer) ·
  [vanilla](https://www.npmjs.com/package/openteams-vanilla-viewer).
- [OneDrive links](https://christophervr.github.io/ooxml/teams/onedrive) in a workspace.
- Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), next to the Word, Excel,
  PowerPoint and Visio viewers.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE); see
[NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/NOTICE).

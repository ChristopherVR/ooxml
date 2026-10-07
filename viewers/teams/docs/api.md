# Element API and events

Every binding wraps the same `<teams-app>` custom element. The bindings forward these properties and re-emit these events in their framework's idiom (see the [framework guides](/frameworks/react)).

## Attributes and properties

| Property        | Attribute       | Description                                                                                                                  |
| --------------- | --------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `workspaceId`   | `workspace-id`  | The shared room: 1 to 128 characters of letters, digits, `-` and `_`.                                                        |
| `userName`      | `user-name`     | Display name.                                                                                                                |
| `userId`        | `user-id`       | Stable id; defaults to one remembered in this browser.                                                                       |
| `config`        | `server-config` | `{ mode: 'local' \| 'server', syncUrl?, signalingUrl?, iceServers, iceTransportPolicy?, token? }`. The attribute takes JSON. |
| `uploadFile`    |                 | Where attachments go; defaults to the server's `/files` endpoint.                                                            |
| `openers`       |                 | Per-Office-type handlers (`docx`, `xlsx`, `pptx`, `vsdx`).                                                                   |
| `embeds`        |                 | Per-file-kind callbacks returning a sandbox-compatible viewer page URL; useful for PowerPoint.                              |
| `client` (read) |                 | The core client behind the element (state in, actions out).                                                                  |

Leave `config` out and the element offers a settings dialog (remembered in this browser) or runs in local mode.

## Content previews

Opening an attachment displays a workspace preview by default. Word and Excel use
the native editor elements, initially read-only; Visio uses its viewer. Markdown has a safe
subset renderer; text files show literal text; HTML and websites use sandboxed
frames. `openers` and the cancelable `teams-open-file` event still take precedence.

PowerPoint requires an embedding page supplied by the host:

```ts
const embeds = {
	pptx: ({ url }: { url: string | undefined }) =>
		`https://viewer.example.com/embed?source=${encodeURIComponent(url ?? '')}`,
};
```

Pass `embeds` to any component binding. The page must understand the file URL and
support a sandbox with scripts/forms but no same-origin permission. File fetches
omit credentials; cross-origin file servers must allow CORS. Without a configured
PowerPoint viewer, the pane reports that limitation and offers external open.

The element's `previewContent({ attachment, url })` method opens host-provided
content locally. For a website use `attachment: { name: 'Project site', kind:
'other', mime: 'text/html' }`. This does not create a shared channel tab. The
Files view also offers **Preview website**. Sites that refuse framing can be
opened externally.

## Shared channel tabs and Excel copies

Use **Pin as tab** beside an uploaded file, or **Add tab** to add a named website.
Definitions are shared; each user selects their own tab. The creator can rename
or remove it through client-side ownership checks. Removing a tab elsewhere keeps
an already open local pane available until it is closed.

Excel offers local editing and **Save copy to channel** when file storage is
configured. Each save uploads a uniquely named copy and posts it in the originating
channel, preserving the original. Failed saves retain edits; navigation confirms
discarding unsaved work and waits for an in-progress save. This is not coediting
or versioned write-back. Hosts forcibly replacing the workspace or element must
handle their own unsaved-work confirmation.

Raw client actions:

```ts
const tab = teams.addTab('Budget', { type: 'file', attachment });
teams.addTab('Project site', { type: 'website', url: 'https://example.com' });
teams.renameTab(tab!.id, 'Budget review');
teams.removeTab(tab!.id);
await teams.saveFileCopy(channelId, file);
```

`TeamsState.tabs` contains the selected channel's definitions;
`canUploadFiles` indicates configured storage. `OpenFileDetail.channelId` records
the channel captured when opening. `previewContent` also accepts this field when
a host wants to enable channel save copies. Ownership checks require server-side
authorization before they can be treated as access control.

## Events

| Event                 | Meaning                                                                             |
| --------------------- | ----------------------------------------------------------------------------------- |
| `teams-ready`         | The element is ready.                                                               |
| `teams-open-file`     | A user opened an attached file. Cancelable: `preventDefault()` to open it yourself. |
| `teams-config-change` | The user saved new server settings in the settings dialog.                          |

Binding spellings: `onReady`, `onOpenFile`, `onConfigChange` (React, Solid, Svelte, vanilla); `@ready`, `@open-file`, `@config-change` (Vue); `(ready)`, `(openFile)`, `(configChange)` (Angular).

## The raw client

`createTeams(options)` (the raw hook in every binding) returns the core client:

```ts
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
teams.on('notice', (notice) => show(notice)); // errors a user should see
```

`getState()` is one immutable `TeamsState`: connection `status`, `channels` with unread counts, the selected channel's `messages` and `files`, `people` (presence), `typing`, the composer state, search results and `call`. Actions include `select`, `createChannel`, `send`, `startReply`, `startEdit`, `deleteMessage`, `toggleReaction`, `notifyTyping`, `setAvailability`, `search`, `openCall`, `joinCall`, `leaveCall`, `toggleMic`, `toggleCamera`, `toggleScreenShare`, `toggleHand` and `destroy`. See [Architecture](/architecture) for how the layers fit.

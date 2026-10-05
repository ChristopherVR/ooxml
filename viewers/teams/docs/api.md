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
| `client` (read) |                 | The core client behind the element (state in, actions out).                                                                  |

Leave `config` out and the element offers a settings dialog (remembered in this browser) or runs in local mode.

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

# The `teams` area (`ooxml-core/teams`)

The logic of an open-source, bring-your-own-server team workspace that sits beside Word, Excel,
PowerPoint and Visio: channels and chat, presence, calls (WebRTC) and links to Office files. It is
**new code** (nothing moved here), built on the `collab` area. DOM-free: storage, `fetch`, media
devices, `RTCPeerConnection` and `WebSocket` are all injectable.

The UI is split the same way as every other product:

| Layer                | Where                                                                                          | Owns                                                                   |
| -------------------- | ---------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------- |
| Logic                | `src/core/teams/` (this area)                                                                  | model, chat CRDT, signaling, calls, server config, client store        |
| Visual primitives    | `src/ui/src/teams/` (`ooxml-ui/teams`)                                                         | `office-ui-avatar`, `-app-rail`, `-channel-list`, `-chat-list`, ...    |
| App, bindings, demos | `viewers/teams` in this repository (the standalone `teams-viewer` repository only points here) | `<teams-app>`, React/Vue/Solid/Svelte/Angular/vanilla bindings, server |

## The client: one state, plain actions

```ts
import { createTeamsClient } from 'ooxml-core/teams';

const teams = createTeamsClient({
	workspaceId: 'acme', // 1-128 chars of [A-Za-z0-9_-]; also the sync room
	user: { id: 'u1', name: 'Ada' },
	config, // see "Bring your own server"
	storage: localStorage, // optional: offline snapshot of the shared document and read markers
});

const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
```

`getState()` returns one immutable `TeamsState` snapshot: connection `status`, `channels` (with
unread counts and a `live` flag), the selected channel with its `messages`, `files` and `pinned`
messages, `people` (presence), your own `presence` (chosen status, message, expiry, idleness), `typing`, the composer's `replyingTo` / `editing`, `searchResults`, `allFiles` and
`call` (`null`, or a phase of `prejoin` / `joining` / `connected` with participants and a camera
preview). Every button is an action: `select`, `createChannel`, `send` (with file uploads),
`startReply`, `startEdit`, `deleteMessage`, `toggleReaction`, `pinMessage`, `unpinMessage`,
`notifyTyping`, `setAvailability` (with an optional reset-after duration), `setStatusMessage`,
`setIdleMinutes`, `noteActivity`, `search`, `openCall`, `setPrejoin`, `joinCall`, `leaveCall`, `toggleMic`, `toggleCamera`,
`toggleScreenShare`, `toggleHand`, `markActivityRead`, `markAllActivityRead`, `setNotificationSettings`, the Chat actions (`openDirectChat`, `createGroupChat`, `renameChat`, `addToChat`, `removeFromChat`, `leaveChat`), the Calendar actions (`createMeeting`, `updateMeeting`, `cancelMeeting`, `cancelOccurrence`, `editOccurrence`, `meetingIcs`, `joinMeeting`), `fileUrl` (async: a short-lived signed link, never the token), `destroy`. State also carries
`chats` (your 1:1 and group chats, recent first, with unread, preview and presence), `contacts` (the workspace roster with presence), `meetings` (in your calendar) and `viewerZone`, `mentionTargets` (who the @ picker offers), `activity` / `activityUnread` (the Activity feed) and
`notificationSettings`. Errors a user should see arrive as `on('notice', ...)`, and
`on('notification', item)` fires for each new activity item that the notification settings and
Do not disturb allow to be announced (the host shows the toast, desktop notification and sound).

This is the shape every binding exposes as a raw hook (`useTeams` in React and Vue,
`createTeamsClient` in Solid, `teamsStore` in Svelte, `TeamsService` in Angular), so an app can
render its own UI without `<teams-app>`.

Lower layers, each usable alone: `createTeamsWorkspace` (collab session plus chat store plus
calls), `createChatStore(doc, user)`, `createCallSession(...)`, the signaling channels, and the
pure projections in `view.ts` (`channelViews`, `filesOf`, `searchMessages`, ...).

## Chat

Channels live in one `Y.Map`, messages in one `Y.Map` per channel, reactions in a flat map keyed
`message|emoji|user`, so two people reacting at once never overwrite each other. Edits and deletes
are author-only (a delete is a soft tombstone so threads keep their shape). Everything a peer wrote
is **re-validated when read back** (`model.ts`): markup is stripped, control characters removed,
lengths capped, attachment URLs must be `http(s)` or a same-origin path, reactions must be one short
token. The UI renders text with `textContent` semantics only.

Files are links, never inlined bytes: `send({ files })` uploads through `uploadFile` (or the
configured server's `/files` endpoint) and posts the returned URL. `detectOfficeKind` maps a name to
`docx` / `xlsx` / `pptx` / `vsdx` so a host can open it in the matching viewer.

### Chats

`chats.ts` holds the model: a chat is `direct` (1:1, id derived from the two user ids) or `group`, with a member list. The directory (`teams:chats` in the workspace document) lists who is in which chat so every client can find its own, and `roster.ts` lists everyone who opened the workspace. A chat's messages live in a Y.Doc of their own, synced in a room named `chatRoomId(workspaceId, chatId)`; `chat-rooms.ts` composes the channel store and those documents behind one `ChatStore`, so threads, pins, reactions, mentions, drafts and the activity feed work unchanged in a chat. The room carries a `teams:chat-members` map (kept equal to the directory by every member) that the reference server admits sockets by, and the client never opens, reads or writes a chat the person is not in. Add, remove and leave write a system message (`Message.system`). `chat-view.ts` projects the chat list and the contacts. Honest limits are in `docs/teams-parity.md`.

### Calendar and meetings

`meetings.ts` stores meetings in the workspace document: title, details, attendees, an optional channel and meeting chat, start and end as UTC milliseconds, repeat (none, daily, weekly) with an end (a count or a date) and per-occurrence exceptions (cancelled or changed, keyed by the occurrence's original start; `meeting-exceptions.ts`, `recurrence.ts`) and the organizer's IANA zone. `time-zone.ts` does the zone arithmetic with `Intl` (wall time to instant, day starts and steps that keep the local time across daylight saving); `calendar.ts` expands occurrences, builds the day, week and work-week views, lays overlapping meetings out in columns and finds due reminders; `ics.ts` writes an iCalendar file in the organizer's zone (TZID, RRULE with COUNT or UNTIL, EXDATE, RECURRENCE-ID overrides) with a VTIMEZONE block that `ics-timezone.ts` derives from `Intl` (two yearly rules where the zone repeats them, otherwise each change written out for ten years). `meeting-actions.ts` ties a meeting to its conversation: a group chat with a meeting card for attendees, or a card in the channel for a channel meeting. `meeting-reminders.ts` announces each occurrence five minutes ahead once per device through the same `shouldNotify` gate as Activity (Off and Do not disturb silence it). `joinMeeting` opens the pre-join screen for a call keyed to the meeting.

### Mentions

`mentions.ts` holds the pure rules: `resolveMentions(text, targets)` finds `@Name` runs (longest name wins, word boundaries, case-insensitive), `sanitizeMentions` re-validates stored ranges against the text, `mentionQueryAt` / `filterMentionTargets` / `insertMention` drive the composer's picker and `mentionSegments` splits text for rendering. A message row stores `mentions` beside `text`; rows from older clients simply have none. `TeamsState.mentionTargets` lists who the picker offers (people who wrote in the channel or are online, then `@channel` and `@team`) and `send({ text, mentions })` takes the picked people.

### Activity and notifications

`activity.ts` derives the feed with `deriveActivity` from the data every client already shares:
mentions of you, replies to your messages or followed threads, reactions to your messages (reaction
rows now store the time they were added) and pins by others, newest first, capped at 200.
`filterActivity` and `describeActivity` back the filter pills and the row text. Read state
(`activity-state.ts`) is personal and local to the device: a watermark plus per-item overrides,
starting at "now" on a first run. `notifications.ts` holds the preferences (all activity, mentions
only or off, and a sound flag) and `shouldNotify`, which also stays silent in Do not disturb. The
core only decides and emits; showing a toast, a browser `Notification` or a sound is the UI's job.

## Shared assistant chats

`createSharedChat({ roomId, user, config })` joins one assistant conversation that lives in its own
room (`sharedChatRoomId()` makes an unguessable `aiaichat-<32 hex>` id), over the same sync transport
the workspace uses (`createSyncTransport`). `store.put` adds or replaces a message by id (unchanged
rows are not written again), `store.truncateFrom` removes a message and everything after it, and
`store.messages()` re-validates every row: role, length caps, author name and a JSON-only `data`
field the host keeps its own fields in. Presence carries what each person is doing (`idle`,
`typing`, `asking`). The suite uses it for the chats people share (see `docs/suite-ai-config.md`).

## Calls

`createCallSession` is a full-mesh set of peer connections using the "perfect negotiation" pattern
(the side with the smaller id is polite, so glare never deadlocks). Audio and video ride two
permanent transceivers whose tracks are swapped with `replaceTrack`, so muting, turning the camera
off and sharing the screen never renegotiate. Mesh needs no media server and is honest up to about a
dozen people (`maxParticipants`, default 12); an SFU would plug in behind the same `CallSession`
surface. `iceTransportPolicy: 'relay'` forces TURN.

## Bring your own server

`TeamsServerConfig` (validated by `parseServerConfig`, which reports every problem and never
throws):

```ts
{ mode: 'local' | 'server', syncUrl?, signalingUrl?, iceServers: [...], iceTransportPolicy?, token?, sessionToken? }
```

- **`local`**: no server. Tabs of one browser share chat and calls over `BroadcastChannel`.
- **`server`**: your server. Any server that honours these three contracts works; the reference
  implementation is `viewers/teams/server/index.mjs` (about 300 lines, Node and `ws`).

### Server contract

1. **Sync**: `ws(s)://host/<base>/<room>` speaks the y-websocket wire format (sync step 1/2/update,
   awareness, query-awareness; see `collab/transport-provider.ts`). A stock y-websocket server works.
   A server that keeps a `Y.Doc` per room and persists it gives history to late joiners.
2. **Signaling**: `ws(s)://host/<base>/<room>` forwards each JSON text message to the **other**
   sockets of the room, and sends `{ "type": "bye", "from": <id> }` to them when a socket closes. A
   socket speaks for one peer id (the first `from` it uses). It does not need to understand the
   messages (`signaling.ts` defines them).
3. **Files** (optional): `POST|GET <origin>/files/<workspace>/<name>`, with the token in an
   `Authorization: Bearer` header. To open a file in a browser tab the client calls
   `POST <origin>/files/link/<workspace>/<name>` (same header) and gets `{ url }` back: a link
   signed for that one file, valid for a few minutes (`?sig=...&exp=...`), so the token never
   appears in a URL.

Private chats add one rule to the sync contract: a room named `chat-<16 hex>-<chat id>` admits a socket only when its `?user=<id>` is in the room's `teams:chat-members` map (a Yjs map of user id to `true`). A room that has no members yet is created by its first socket, which passes `?members=<id>,<id>` and must be one of them; a 1:1 room (`dm-<hash>`) must list exactly the two people the id was derived from. The server closes the sockets of people removed from the map. Identity is a claim unless the server has accounts (see Authentication below).

Both sockets receive `?token=<token>` when a token is configured (a browser cannot set headers on
a WebSocket); enforce it, and an origin allowlist, on the server. Validate room names (`^[\w-]{1,128}$`) and cap message sizes.

### Authentication

Optional, and off by default. A server with accounts exposes `POST <origin>/auth/session` (body
`{ user, credential }`, the workspace token in `Authorization` when one is set) and answers
`{ session, user, expiresAt }`: a signed, expiring token that binds one user id.
`requestTeamsSession(config, { userId, credential })` calls it; pass the result as
`config.sessionToken` and both sockets send `?session=<token>`. The server then takes the user id
from the session and ignores the claimed `?user=`, so holding the shared token no longer lets a
client join another person's chats. With accounts the server refuses sockets that have no valid
session (401).

A server without accounts (the zero-config dev default) signs a session for any claimed id and keeps
trusting the claim: that is **not secure**. The reference server's `authenticate` hook, `users`
table and `TEAMS_USERS_FILE` are described in its README. Not covered: the session rides in the
WebSocket query string like the token, presence and awareness remain advisory, sessions cannot be
revoked before they expire, and the UI has no sign-in form yet (a host calls `requestTeamsSession`
itself).

### STUN and TURN

`iceServers` carries STUN and TURN URLs. Public STUN is enough on open networks and never enough
behind symmetric NAT: run your own TURN (for example coturn with `use-auth-secret`) and give the
clients credentials. TURN entries without credentials are rejected by `parseServerConfig`.

## Honest limits

- **No end-to-end encryption.** Chat travels as Yjs updates through your server, which can read
  them. Media is DTLS-SRTP between peers (or to your TURN relay, which cannot decrypt it).
- **Authorization is the server's job.** The `role` and author checks here are client-side; a
  hostile client can write anything the server accepts. Treat the server as the trust boundary.
- **Mesh calls scale to a handful of people**, not a webinar.
- A sync server that never compacts will grow with history; the reference server stores one
  snapshot per room.
- Presence is advisory and unauthenticated, like any Yjs awareness. The wire carries the original
  `availability` (available, busy, away) next to the richer `status` (also Do not disturb, Be right
  back, Appear offline), so older clients still read presence; a status message rides along.
- Private chat rooms keep non-members from receiving a chat, and with accounts enabled identity comes from the signed session; without them it is the `user` the client claims, so anyone with the shared token can claim another person's id. In both cases the chat directory (who talks to whom) and attachments are visible to the workspace. Meeting rows are in the shared document too: the attendee list scopes a calendar, not who could read the row.
- Pins are shared rows in a per-channel Yjs map; any member may pin or unpin, as in Teams. There is
  no role model, so this is not an enforced permission.
- No claim of Microsoft Teams feature parity beyond what the UI actually implements (channels,
  posts with replies and thread panes, reactions, pinned messages, files, search, presence with
  status messages, @mentions, an activity feed, meetings with screen share and raise hand).
  Notifications are in-page only (toast, browser Notification while the tab is open, sound): no
  push, email or service worker. 1:1 and group chats and a calendar with
  meetings, repeats, reminders and .ics export exist (see `docs/teams-parity.md` for their limits).
  Not implemented: recording, background blur, live captions.

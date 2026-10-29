# Limitations

OpenTeams is an early implementation. This page lists what it does not do, so you can decide
before you rely on it.

## Not Microsoft Teams

OpenTeams is not Microsoft Teams, is not affiliated with Microsoft and makes no claim of feature
parity. It implements:

- channels, posts with replies and thread panes, reactions, author-only edit and delete, search and unread counts;
- 1:1 and group chats, each in its own room that only its members can join;
- an Activity feed and in-page notifications (a toast, a browser notification while the tab is open, a sound);
- a calendar with meetings, repeats, reminders and `.ics` export;
- presence: who is online, typing, a status you set (available, busy, do not disturb, be right back, away, offline) with an optional reset time, a status message and automatic away when idle;
- pinned messages shown in a bar at the top of the channel (any member can pin or unpin);
- meetings in a channel: audio, video, screen share and raise hand, with a pre-join screen;
- files shared in the conversation as links, with Word, Excel, PowerPoint and Visio files
  recognised by name.

It does **not** implement recording, background blur, live captions, push or email notifications
(notifications work only while the app is open), invitations outside the workspace or RSVP.

## Content workspace

Word previews read-only; Excel offers local editing and saving new channel copies
when storage is configured; Visio uses its existing viewer.
There is no coediting, overwrite or versioned save-back. PowerPoint needs a
host-configured embedded viewer page. Missing integrations are reported in the pane.
Text reads are limited to 2 MiB and native Office reads to 32 MiB. The interface is available in
English, German, Spanish, French and Simplified Chinese (the `locale` prop); text people wrote is
never translated.

Files views support direct upload, blank Excel creation and metadata search when
storage is configured. Uploads show a busy state and allow retry, with no byte
progress or cancellation. A failed batch can leave files that were uploaded before the failure and are not
referenced by any message; the reference server removes its partial uploads and lets the uploader or a
moderator delete the rest (`DELETE /files/<scope>/<name>`), but does not remove them on its own.
The reference server limits private chat files to the chat's members and counts storage per
workspace; channel files are readable by every signed-in user of the server. Folders, document
content indexing and file version history are not implemented.

Markdown supports headings, bullets, quotes, fenced code, basic emphasis and web
links, flat task lists, top-level pipe tables and file-relative web links. It does not
implement full CommonMark/GFM, nested structures or reference links. Raw HTML is displayed as text.

Websites and HTML attachments use a sandbox without same-origin, popup or top
navigation permission. Some sites block embedding or require permissions this
sandbox does not grant. External open remains available. File and website tab
definitions can be shared in channels; the embedded app platform, authentication
and permissions of Microsoft Teams are not implemented. Tab creator checks run
in the client and require server authorization to become enforceable.

## Security and privacy

- **No end-to-end encryption.** Chat travels as Yjs updates through your server, which can read
  and stores them. Uploaded files are stored as they are. Call media is DTLS-SRTP between peers (a
  TURN relay cannot decrypt it).
- **Authorization is the server's job.** Author and role checks run in the client; a hostile
  client can write anything the server accepts.
- **Authentication is required by default.** The reference server refuses to start unless it has
  accounts, an `authenticate` hook or `authorizeSync`; `--insecure-dev` opts into an open development
  mode with a warning. With accounts it signs sessions for configured users (or your `authenticate`
  hook) and takes the user id from the session, and `authorizeSync` can decide access per room. A
  shared `TEAMS_TOKEN` alone is one secret stored in the browser, not per-user authentication, and
  user ids are claims. There is no SSO, password reset or team and channel access control. See [what the reference server does not do](/server#what-the-reference-server-does-not-do).
- **Presence is advisory and unauthenticated**, like any Yjs awareness: a client can claim any
  name.
- Everything a peer sends is re-validated when read back and rendered as text, never as markup.

## Scale

- **Mesh calls** connect every participant to every other one. That needs no media server and
  works for a handful of people (the default cap is 12), not a webinar.
- **History grows.** The reference server keeps one snapshot per room and never compacts it.
- **One server process.** The reference server keeps rooms in memory and does not scale
  horizontally.

## Local mode and the live demos

`mode: 'local'` uses `BroadcastChannel`, so only tabs of the same browser on the same site see
each other. Nothing is sent to any server, files are shared by name only (there is nowhere to upload
them), and the shared document is kept in that browser's `localStorage`. The
[live demos](/demos) run this way because GitHub Pages cannot host a server.

## Calls across networks

Public STUN works on open networks. Behind symmetric NAT or strict firewalls you need your own TURN
server; without it, some calls will not connect. Cameras and microphones need `localhost` or
HTTPS.

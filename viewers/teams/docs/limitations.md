# Limitations

OpenTeams is an early implementation. This page lists what it does not do, so you can decide
before you rely on it.

## Not Microsoft Teams

OpenTeams is not Microsoft Teams, is not affiliated with Microsoft and makes no claim of feature
parity. It implements:

- channels, posts with replies, reactions, author-only edit and delete, search and unread counts;
- presence: who is online, typing, and an availability you set (available, busy, away);
- meetings in a channel: audio, video, screen share and raise hand, with a pre-join screen;
- files shared in the conversation as links, with Word, Excel, PowerPoint and Visio files
  recognised by name.

It does **not** implement 1:1 or group chats outside channels, threads as separate panes,
notifications, a calendar, recording, background blur or live captions.

## Content workspace

Word and Excel preview in native read-only editors; Visio uses its existing viewer.
There is no shared editing or save-back to file storage. PowerPoint needs a
host-configured embedded viewer page. Missing integrations are reported in the pane.
Text reads are limited to 2 MiB and native Office reads to 32 MiB.

Markdown supports headings, bullets, quotes, fenced code, basic emphasis and web
links. It does not yet implement full CommonMark/GFM, tables, task lists, nested
structures or relative links. Raw HTML is displayed as text.

Websites and HTML attachments use a sandbox without same-origin, popup or top
navigation permission. Some sites block embedding or require permissions this
sandbox does not grant. External open remains available. This is a local preview,
not a shared configurable channel tab or a Microsoft Teams app integration.

## Security and privacy

- **No end-to-end encryption.** Chat travels as Yjs updates through your server, which can read
  and stores them. Uploaded files are stored as they are. Call media is DTLS-SRTP between peers (a
  TURN relay cannot decrypt it).
- **Authorization is the server's job.** Author and role checks run in the client; a hostile
  client can write anything the server accepts.
- **The reference server has no user accounts.** Its token is one shared secret stored in the
  browser, not per-user authentication. See [what the reference server does not do](/server#what-the-reference-server-does-not-do).
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

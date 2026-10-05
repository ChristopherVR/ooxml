# The live demos

Two demos are published with this site:

| Demo                                              | What it shows                                                                                         | Source         |
| ------------------------------------------------- | ----------------------------------------------------------------------------------------------------- | -------------- |
| [Vanilla](/demo/){target="_self"}                 | The whole app: one `<teams-app>` element.                                                              | `demos/vanilla` |
| [React](/demo-react/){target="_self"}             | `<Teams />` on the left and a hand-written panel over the `useTeams()` hook on the right.              | `demos/react`   |

## They run without a server

GitHub Pages only serves static files, so there is no server behind these pages and nothing is
faked to look like one. The demos are built with `VITE_TEAMS_STATIC=1`, which starts them in the
core's **local mode**:

- chat, presence and call signaling travel between tabs of **your browser only**, over
  `BroadcastChannel`;
- the shared document is kept in your browser's `localStorage`, so it survives a reload;
- files are shared by name only, because there is nowhere to upload them;
- calls connect tabs of the same browser (your camera and microphone, with your permission).

Each demo shows a notice saying so, with a link that opens a second tab as another person (the
`?name=` parameter). Type in one tab and the message appears in the other.

If you saved your own server in the vanilla demo's **Server** settings, it uses that instead; clear
it there (or clear the site's storage) to return to local mode.

## Use your own server

Real use needs a server: run `npx openteams-server` and see [Bring your own server](/server). From
source, `bun run dev` starts the reference server and both demos against it (see
[Getting started](/getting-started#run-it-from-source)).

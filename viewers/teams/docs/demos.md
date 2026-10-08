# The live demos

One demo per binding is published with this site. All of them run the same workspace and take the same query parameters (`?name=Ada&room=acme`).

| Demo                                                                                         | What it shows                                                                             | Source          |
| -------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- | --------------- |
| [Vanilla](/demo/){target="_self"} (also at [/demo-vanilla/](/demo-vanilla/){target="_self"}) | The whole app: one `<teams-app>` mounted by `mountTeams`.                                 | `demos/vanilla` |
| [React](/demo-react/){target="_self"}                                                        | `<Teams />` on the left and a hand-written panel over the `useTeams()` hook on the right. | `demos/react`   |
| [Vue](/demo-vue/){target="_self"}                                                            | The `<Teams>` component of `openteams-vue-viewer`.                                        | `demos/vue`     |
| [Angular](/demo-angular/){target="_self"}                                                    | `<teams-workspace>` of `openteams-angular-viewer`.                                        | `demos/angular` |
| [Svelte](/demo-svelte/){target="_self"}                                                      | `Teams.svelte` of `openteams-svelte-viewer`.                                              | `demos/svelte`  |
| [Solid](/demo-solid/){target="_self"}                                                        | `Teams()` of `openteams-solid-viewer`.                                                    | `demos/solid`   |

Because they share one origin and one local room name, a demo in one framework can talk to a demo in another: open `/demo-vue/?name=Ada&room=x` and `/demo-svelte/?name=Bob&room=x` in two tabs.

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

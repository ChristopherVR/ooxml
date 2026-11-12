# Die Live-Demos

Zu dieser Website wird eine Demo pro Binding veröffentlicht. Alle laufen im selben Arbeitsbereich und nehmen dieselben
Abfrageparameter entgegen (`?name=Ada&room=acme`).

| Demo                                                                                            | Was gezeigt wird                                                                        | Quelle                |
| ----------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- | --------------------- |
| [Vanilla](/demo/){target="_self"} (auch unter [/demo-vanilla/](/demo-vanilla/){target="_self"}) | Die ganze App: ein `<teams-app>`, eingebunden von `mountTeams`.                         | `demos/teams/vanilla` |
| [React](/demo-react/){target="_self"}                                                           | `<Teams />` links und ein selbst geschriebenes Panel über dem Hook `useTeams()` rechts. | `demos/teams/react`   |
| [Vue](/demo-vue/){target="_self"}                                                               | Die Komponente `<Teams>` aus `openteams-vue-viewer`.                                    | `demos/teams/vue`     |
| [Angular](/demo-angular/){target="_self"}                                                       | `<teams-workspace>` aus `openteams-angular-viewer`.                                     | `demos/teams/angular` |
| [Svelte](/demo-svelte/){target="_self"}                                                         | `Teams.svelte` aus `openteams-svelte-viewer`.                                           | `demos/teams/svelte`  |
| [Solid](/demo-solid/){target="_self"}                                                           | `Teams()` aus `openteams-solid-viewer`.                                                 | `demos/teams/solid`   |

Weil alle Demos denselben Ursprung und denselben lokalen Raumnamen teilen, kann eine Demo in einem Framework mit einer
Demo in einem anderen sprechen: Öffnen Sie `/demo-vue/?name=Ada&room=x` und `/demo-svelte/?name=Bob&room=x` in zwei Tabs.

## Sie laufen ohne Server {#they-run-without-a-server}

GitHub Pages liefert nur statische Dateien. Hinter diesen Seiten gibt es daher keinen Server, und nichts wird so
vorgetäuscht, als gäbe es einen. Die Demos werden mit `VITE_TEAMS_STATIC=1` gebaut. Dadurch starten sie im **lokalen
Modus** des Core:

- Chat, Präsenz und Anrufsignalisierung laufen nur zwischen Tabs **Ihres Browsers**, über `BroadcastChannel`;
- das gemeinsame Dokument wird im `localStorage` Ihres Browsers gehalten und überlebt so einen Neuladen;
- Dateien werden nur über ihren Namen geteilt, weil es keinen Ort gibt, an den sie hochgeladen werden könnten;
- Anrufe verbinden Tabs desselben Browsers (Ihre Kamera und Ihr Mikrofon, mit Ihrer Erlaubnis).

Jede Demo zeigt einen Hinweis dazu, mit einem Link, der einen zweiten Tab als eine andere Person öffnet (über den
Parameter `?name=`). Schreiben Sie in einem Tab, erscheint die Nachricht im anderen.

Haben Sie in den Einstellungen **Server** der Vanilla-Demo einen eigenen Server gespeichert, nutzt die Demo diesen
stattdessen. Löschen Sie ihn dort (oder löschen Sie den Speicher der Website), um zum lokalen Modus zurückzukehren.

## Eigenen Server verwenden {#use-your-own-server}

Für den echten Einsatz brauchen Sie einen Server: Starten Sie `npx openteams-server --insecure-dev` (nur für die lokale
Entwicklung) und lesen Sie [Eigenen Server mitbringen](/de/server). Aus dem Quellcode startet `bun run dev` den
Referenzserver und beide Demos, die ihn nutzen (siehe [Erste Schritte](/de/getting-started#run-it-from-source)).

# Zusammenarbeit

Zusammenarbeit ist das, was OpenTeams ausmacht: Jeder Kanal, jede Nachricht, jede Reaktion und jedes Präsenzsignal ist gemeinsamer Zustand, den mehrere Personen gleichzeitig bearbeiten. Diese Seite erklärt, wie das aufgebaut ist, was die Live-Demo zeigt und was nicht.

## So funktioniert es

- **Das gemeinsame Dokument ist ein Yjs-CRDT.** Kanäle liegen in einer `Y.Map`, Nachrichten in je einer `Y.Map` pro Kanal und Reaktionen in einer flachen Map, geschlüsselt nach Nachricht, Emoji und Benutzer, sodass gleichzeitige Bearbeitungen und Reaktionen zusammengeführt werden, statt sich gegenseitig zu überschreiben. Die Logik liegt in `ooxml-core/teams`, aufgebaut auf dem formatneutralen Bereich `ooxml-core/collab`.
- **Präsenz ist Yjs-Awareness**: wer online ist, wer gerade tippt und eine Verfügbarkeit, die Sie selbst setzen (verfügbar, beschäftigt, abwesend). Sie ist beratend und nicht authentifiziert, wie jede Yjs-Awareness.
- **Anrufe sind WebRTC.** Die Signalisierung läuft über Ihren Server (oder `BroadcastChannel` im lokalen Modus); Medien gehen Peer-to-Peer in einem vollständigen Mesh, daher eignen sich Anrufe für eine Handvoll Personen.
- **Der Transport ist austauschbar.** Es gibt zwei Modi: `server` (Ihr Server spricht das y-websocket-Sync-Protokoll, leitet die Anruf-Signalisierung weiter und speichert optional Dateien) und `local` (Tabs eines Browsers teilen alles über `BroadcastChannel`).

## Live-Demo: zwei Personen, ein Raum, beliebiges Framework

Die beiden Bereiche unten sind zwei getrennte Demo-Apps in einem lokalen Raum. Ada nutzt das Vanilla-`<teams-app>` und Bob das React-Binding, und doch teilen sie einen Arbeitsbereich: Ein Raum ist nur ein Name (`?room=`), und jedes Binding spricht dasselbe Core-Protokoll. Sie laufen im lokalen Modus und sprechen daher nur über Ihren Browser miteinander. Schreiben Sie als Ada eine Nachricht und sehen Sie zu, wie sie bei Bob erscheint.

Auf der Startseite hat bei der [Live-Demo](/de/#live-demo) jeder Bereich eine eigene Framework-Auswahl, sodass Sie zwei beliebige aus Vanilla, React, Vue, Angular, Svelte und Solid kombinieren können.

<iframe
	src="/ooxml/teams/demo/?name=Ada&room=docs-collab"
	title="OpenTeams demo as Ada"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/teams/demo-react/?name=Bob&room=docs-collab"
	title="OpenTeams React demo as Bob"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

Dasselbe funktioniert über echte Tabs hinweg, in jeder Kombination: Öffnen Sie [die Vanilla-Demo als Ada](/demo/?name=Ada&room=docs-collab){target="_self"} und zum Beispiel [die Vue-Demo als Bob](/demo-vue/?name=Bob&room=docs-collab){target="_self"}. Jede Demo nimmt dieselben Parameter `name` und `room` entgegen; im lokalen Modus können sich nur Frames und Tabs desselben Browsers treffen, für Personen auf verschiedenen Rechnern richten Sie daher jedes Binding auf Ihren eigenen Server (siehe unten).

## Was unterstützt wird und was nicht

| Unterstützt                                                                 | Nicht unterstützt                                                                                                                     |
| --------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| Gleichzeitiger Chat, Antworten und Reaktionen, per Yjs-CRDT zusammengeführt | Ende-zu-Ende-Verschlüsselung: Ihr Server kann alles lesen und speichern                                                               |
| Präsenz und Tippanzeigen                                                    | Zugriffssteuerung pro Team und SSO; die Identität ist eine Behauptung, solange Sie die optionalen Konten des Servers nicht aktivieren |
| Zusammenarbeit im selben Browser ohne Server (lokaler Modus, wie oben)      | Zusammenarbeit zwischen verschiedenen Rechnern ohne einen Server, den Sie betreiben                                                   |
| Rechnerübergreifende Zusammenarbeit über Ihren eigenen Sync-Server          | Ein gehosteter Dienst: Es gibt keinen, und die Pages-Demos haben keinen Server dahinter                                               |
| Anrufe mit Audio, Video, Bildschirmfreigabe und Handheben (Mesh, ~12)       | Anrufe in Webinar-Größe, Aufzeichnung, Push-Benachrichtigungen                                                                        |

## Über Rechner hinweg

Betreiben Sie den Referenzserver (`npx openteams-server`) oder einen beliebigen Server, der die drei Verträge einhält (Synchronisierung, Signalisierung, optionale Dateien), und setzen Sie `config.mode` auf `'server'`. Siehe [Bringen Sie Ihren eigenen Server mit](/de/server) und [Bereitstellung](/de/deploy). Das gemeinsame Bearbeiten von Office-Dateien gehört nicht zu OpenTeams: Dateien werden als Links geteilt, und ein Host kann sie im Word-, Excel-, PowerPoint- oder Visio-Viewer öffnen, jeweils mit eigener Zusammenarbeit.

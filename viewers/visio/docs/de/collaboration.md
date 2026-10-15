# Zusammenarbeit {#collaboration}

Datei > Freigeben startet eine Live-Sitzung zwischen Fenstern eines Browsers. Das Teilen über mehrere Geräte hinweg ist noch nicht verfügbar.

::: warning Teilweise umgesetzt, nur im selben Browser
In diesem Build gibt es weder einen gehosteten Dienst noch einen Server-Transport. Nichts verlässt Ihren Browser, und es wird kein Dokument hochgeladen.
:::

## Was es heute gibt {#what-exists-today}

- Eine Live-Sitzung auf Basis von `ooxml-core/collab` (Yjs) über einen `BroadcastChannel`, sodass Fenster und Tabs desselben Browsers eine Zeichnung gemeinsam nutzen können.
- Das gesamte VSDX-Paket wird geteilt, und die Collab-Laufzeit wird erst geladen, wenn das Teilen beginnt.
- Entfernte Änderungen treffen als ein einziger rückgängig machbarer Verlaufsschritt ein, und Mitwirkende erscheinen über das gemeinsame `ooxml-ui`-Präsenzelement.
- Der Sitzungsname besteht aus 1 bis 64 Buchstaben, Ziffern, Bindestrichen oder Unterstrichen. Ein Fenster, das einer Sitzung beitritt, deren Raum bereits eine Zeichnung enthält, übernimmt diese; andernfalls befüllt es den Raum.

## Live-Demo: zwei Fenster, beliebiges Framework {#live-demo-two-windows-any-framework}

Eine Freigabesitzung ist nur ein Name. Beide Bereiche unten sind eigenständige Playgrounds, die mit unterschiedlichen Framework-Bindings gebaut wurden und beim Laden der Sitzung `docs-collab` beitreten: Fenster A ist das Vanilla-Playground und öffnet das Beispiel, Fenster B ist das React-Playground und erhält die Zeichnung aus dem Raum. Bearbeiten Sie den Formtext in einem Bereich und beobachten Sie, wie der andere folgt; Rückgängig macht den entfernten Schritt rückgängig.

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer Fenster A (Vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer Fenster B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Zwei beliebige Frameworks koppeln {#pair-any-two-frameworks}

Jedes Playground akzeptiert `?share=<session>`, das Datei > Freigeben ausfüllt und startet. Nur der Sitzungsname zählt, daher funktioniert jede Kombination: Die [Live-Demo](/de/#live-demo) auf der Startseite bietet für jedes Fenster eine Framework-Auswahl, und in zwei Tabs können Sie zum Beispiel `/demo-vue/?sample=1&share=my-room` und `/demo-angular/?share=my-room` öffnen. Sie können es auch von Hand in zwei beliebigen Fenstern tun: Datei, dann Freigeben, dann derselbe Sitzungsname.

### Warum es nur im selben Browser funktioniert {#why-it-is-same-browser-only}

Die Sitzung nutzt `ooxml-core/collab` (Yjs) über einen `BroadcastChannel`. Ein `BroadcastChannel` verbindet Fenster, Tabs und Frames eines Browserprofils auf einer Origin, unabhängig vom verwendeten Framework, und sonst nichts. Zwei verschiedene Rechner können sich auf diese Weise nicht treffen; dafür ist ein Server-Transport nötig, den die Host-Anwendung bereitstellt, und dieser Build liefert keinen mit.

## Was unterstützt wird und was nicht {#what-is-and-is-not-supported}

| Unterstützt                                                                               | Nicht unterstützt                                                                                               |
| ----------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| Sitzungen im selben Browser über BroadcastChannel, auf Basis von Yjs                      | Teilen über Geräte hinweg: erfordert einen vom Host bereitgestellten Server-Transport, keiner wird mitgeliefert |
| Teilen des gesamten Pakets mit einem rückgängig machbaren Schritt pro entfernter Änderung | Zusammenführen auf Feldebene: gleichzeitige Änderungen werden nicht zusammengeführt, das letzte Paket gewinnt   |
| Mitwirkende werden mit dem gemeinsamen Präsenzelement angezeigt                           | Kommentare, Authentifizierung und Rollen                                                                        |
| Datei > Freigeben als Einstiegspunkt                                                      | Eine dokumentierte Einbettungs-API für eigene Transporte                                                        |

Die Nachweise sind im [Funktionsverzeichnis](/parity) unter „Collaboration and comments“ festgehalten.

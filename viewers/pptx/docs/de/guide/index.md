---
title: Überblick über die Dokumentation
description: Eine Karte der pptx-viewer-Dokumentation - welcher Abschnitt relevant ist, je nachdem ob Sie einen Viewer einbetten, .pptx-Dateien per Code verarbeiten, den Editor benutzen oder mit Agenten automatisieren.
---

# Überblick über die Dokumentation

Die Dokumentation ist nach Zielen gegliedert. Nutzen Sie diese Seite, um den passenden Abschnitt zu finden, und folgen Sie dann der vorgeschlagenen Lesereihenfolge.

## Wählen Sie einen Pfad

| Sie möchten...                                                          | Beginnen Sie mit                                                                                                                                                                                                      | Danach                                                                                                                  |
| ----------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| Einen PowerPoint-Viewer oder -Editor in eine Web-App einbetten          | Der Einstiegsseite Ihres Frameworks: [React](/react/getting-started), [Vue 3](/vue/getting-started), [Angular](/angular/getting-started), [Svelte 5](/svelte/getting-started), [Vanilla JS](/vanilla/getting-started) | Der Props/API-Referenz desselben Abschnitts, dann [Theming](/guide/theming) und [Lokalisierung](/de/guide/localization) |
| `.pptx`-Dateien per Code lesen, bearbeiten oder erzeugen                | [Core: Laden und Parsen](/core/loading) und [die Builder-API](/core/builder)                                                                                                                                          | [Bearbeitung](/core/editing), [Speichern](/core/saving), [das Datenmodell](/de/guide/data-model)                        |
| Präsentationen in Markdown, Bilder, PDF oder SVG umwandeln              | [Markdown-Konverter](/core/converter) und [SVG-Export](/core/svg-export)                                                                                                                                              | Die Export-Seite Ihres Bindings für Browser-Formate                                                                     |
| Teile der Viewer-Oberfläche per Code ausblenden, sperren oder umbelegen | [UI-Anpassung](/guide/customization) (auf Englisch)                                                                                                                                                                   | Rezepte auf dieser Seite: Kiosk-Modus, reduziertes Menüband, festgelegte Einstellungen                                  |
| Mehrere Office-Editoren einbetten und Host-Steuerung festlegen          | [Einbettungsleitfaden](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md)                                                                                                                            | Funktionen, Dateiseiten, Menüs, Bereiche, Speicher und schreibgeschützter Modus                                         |
| Den Editor benutzen (ohne Code zu schreiben)                            | [Benutzerhandbuch](/de/user/)                                                                                                                                                                                         | [Folien bearbeiten](/de/user/editing), [Tastenkürzel](/de/user/shortcuts)                                               |
| KI-Agenten oder Skripte mit Präsentationen arbeiten lassen              | [MCP und Werkzeuge](/packages/mcp)                                                                                                                                                                                    | [Core-CLI](/core/cli)                                                                                                   |
| Die interne Funktionsweise der Bibliothek verstehen                     | [Architektur](/de/guide/architecture)                                                                                                                                                                                 | [OOXML-Konformität](/architecture/openxml-conformance), [Einschränkungen](/de/guide/limitations)                        |

## Neu im Projekt

Wenn noch kein Pfad passt, lesen Sie diese drei Seiten in dieser Reihenfolge:

1. [Was ist pptx-viewer?](/de/guide/introduction): die Paketfamilie und die Rolle jedes Teils.
2. [Installation](/de/guide/installation): welches Paket für Ihren Stack.
3. [Schnellstart](/de/guide/quick-start): vier vollständige Abläufe mit der öffentlichen API.

## Abschnitte im Überblick

| Abschnitt                                                                                                              | Zielgruppe                          | Inhalt                                                                                                    |
| ---------------------------------------------------------------------------------------------------------------------- | ----------------------------------- | --------------------------------------------------------------------------------------------------------- |
| **Entwicklerhandbuch** (dieser Abschnitt)                                                                              | Entwickler                          | Konzepte für alle Pakete: Architektur, Datenmodell, Theming, i18n, Einschränkungen.                       |
| **[Core](/core/)**                                                                                                     | Entwickler                          | Die Headless-Engine: laden, bearbeiten, erstellen, speichern, konvertieren, verschlüsseln. Keine UI.      |
| **[React](/react/)** / **[Vue](/vue/)** / **[Angular](/angular/)** / **[Svelte](/svelte/)** / **[Vanilla](/vanilla/)** | Entwickler                          | Ein Abschnitt pro UI-Binding: Einstieg, Props/API, Theming, Export, Zusammenarbeit.                       |
| **[Benutzerhandbuch](/de/user/)**                                                                                      | Endanwender des Editors             | Anzeigen, Bearbeiten, Präsentieren, Exportieren und Zusammenarbeiten in der Editor-Oberfläche. Kein Code. |
| **[MCP und Werkzeuge](/packages/mcp)**                                                                                 | Autoren von Agenten/Automatisierung | Der MCP-Server, seine 73 Werkzeuge und die CLI.                                                           |
| **[Releases](/releases/)**                                                                                             | Alle                                | Versionshinweise pro Paket.                                                                               |

::: info Sprache
Die Konzeptseiten des Handbuchs sind ins Deutsche übersetzt; die detaillierten Paketreferenzen (Core, React, Vue, Angular, Svelte, Vanilla) sowie Theming und UI-Anpassung sind auf Englisch.
:::

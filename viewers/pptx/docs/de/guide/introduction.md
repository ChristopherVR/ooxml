---
title: Was ist pptx-viewer?
description: Ein allgemeiner Überblick über das TypeScript-Monorepo pptx-viewer zum Analysieren, Bearbeiten, Rendern und Konvertieren von PowerPoint-Dateien mit React, Vue 3, Angular, Svelte 5 und Vanilla JavaScript.
---

# Was ist pptx-viewer?

`pptx-viewer` ist ein TypeScript-Monorepo zum **Analysieren, Bearbeiten, Rendern und Konvertieren** von Microsoft-PowerPoint-Dateien (`.pptx`) - im Browser und in Node.js. Es arbeitet vollständig im Arbeitsspeicher mit dem OpenXML-ZIP-Archiv und ohne native Abhängigkeiten.

`pptx-viewer` deckt den vollständigen Kreislauf ab: eine vorhandene Datei laden, ihr strukturiertes Datenmodell verändern, es rendern und als gültige `.pptx`-Datei zurückspeichern. Dieselbe Core-Engine treibt Komponenten für **React**, **Vue 3**, **Angular** und **Svelte 5** an, dazu einen **Vanilla-JavaScript**-Build für Projekte ohne Framework.

## Was es macht

Das SDK bietet neun Kernfähigkeiten:

1. **Analysieren** von `.pptx` (und `.ppsx` / `.pptm` / `.potx`, dazu das alte Binärformat `.ppt` aus PowerPoint 97-2003) aus einem rohen `ArrayBuffer` in ein strukturiertes [`PptxData`](/de/guide/data-model)-Modell.
2. **Erstellen** von Präsentationen von Grund auf mit einer fließenden Builder-API.
3. **Rendern** von Folien als interaktive React-, Vue-, Angular-, Svelte- oder Vanilla-JS-Komponenten.
4. **Bearbeiten** von Präsentationen programmatisch oder über den integrierten WYSIWYG-Editor.
5. **Speichern** von Änderungen zurück in eine gültige `.pptx`-Datei (roundtrip-sicher).
6. **Konvertieren** von Präsentationen in Markdown mit optionaler Medienextraktion.
7. **Exportieren** von Folien als Bilder (PNG/JPEG), SVG, PDF, GIF oder Video.
8. **Zusammenarbeiten** in Echtzeit über Yjs-CRDT mit Anwesenheitsanzeige.
9. **Verschlüsseln/Entschlüsseln** von passwortgeschützten PPTX-Dateien (AES-128/256).

Die Engine verarbeitet Präsentationselemente, Hunderte von `ST_ShapeType`-Voreinstellungsformen, Diagramme, integrierte Tabellenformatvorlagen, SmartArt, 3D-Modelle, Animationen und Übergänge einschließlich Morph, Designs, Folienmaster, eingebettete Medien, EMF/WMF-Metadateien, OLE-Objekte, digitale Tinte, digitale Signaturen, Verschlüsselung, VBA-Makro-Erhaltung und OOXML-Strict-Konformität.

## Für wen ist es gedacht?

- **Produktteams, die einen Dokument-Viewer einbetten**: zeigen Sie hochgeladene Präsentationen in einer Web-App an, ohne Dateien an einen Konvertierungsdienst zu senden.
- **Teams, die einen Editor bauen**: die Bindings enthalten eine vollständige WYSIWYG-Bearbeitungsoberfläche (Menüband, Inspektor, Rückgängig/Wiederholen, Zusammenarbeit), die Sie funktionsweise aktivieren können.
- **Backend und Automatisierung**: das Core-Paket ist framework-unabhängig und läuft in Node.js, sodass Sie Präsentationen in Skripten, auf Servern und in CI erzeugen, prüfen, vergleichen oder konvertieren können.
- **KI- und Agenten-Pipelines**: `pptx-viewer-mcp` stellt die Engine als MCP-Tool-Aufrufe bereit, und der Markdown-Konverter macht aus Präsentationen LLM-freundlichen Text.

## Rendering-Modell

`pptx-viewer` rastert Folien nicht, weder serverseitig noch auf einem `<canvas>`. Folien werden als **HTML, CSS und SVG im DOM** gerendert:

- Text ist echter HTML-Text: auswählbar, durchsuchbar, übersetzbar und für Screenreader lesbar.
- Formen und Verbinder sind SVG mit berechneten Clip-Pfaden aus der Geometrie-Engine.
- Tabellen sind HTML-`<table>`-Elemente; Diagramme sind Inline-SVG, aufgebaut aus den analysierten Diagrammdaten.
- Zoomen ist eine CSS-Transformation, sodass Folien auf jeder Anzeigedichte bei jeder Skalierung scharf bleiben.
- Da jedes Element ein DOM-Knoten ist, erhält der Editor Auswahl, Ziehgriffe, Inline-Textbearbeitung und Barrierefreiheit gratis.

Rasterung findet nur statt, wenn Sie sie anfordern: Der Raster-Export (PNG/JPEG/PDF/GIF/Video) versucht zuerst eine SVG-`foreignObject`-Erfassung (die Engine des Browsers zeichnet das DOM selbst und erhält `backdrop-filter`, benutzerdefinierte Eigenschaften und 3D-Transformationen), weicht dann auf die erneute Rasterung der vorhandenen Vektor-SVG-Export-Pipeline aus und zeichnet das DOM erst als letzten Ausweg über `html2canvas-pro`.

::: tip Kompromisse
CSS kann nicht jeden PowerPoint-Effekt pixelgenau ausdrücken, und der `html2canvas-pro`-Ausweichpfad (wenn eine Ressource für die SVG-Erfassung nicht eingebettet werden kann) hat eigene Lücken, zum Beispiel die Annäherung von Mischmodi durch Deckkraft. Die vollständige Liste finden Sie unter [Einschränkungen](/de/guide/limitations).
:::

## Die Pakete

Das Monorepo stellt acht veröffentlichte Pakete bereit.

| Paket                            | npm-Name                     | Zweck                                                                                                         |
| -------------------------------- | ---------------------------- | ------------------------------------------------------------------------------------------------------------- |
| [**Core**](/core/)               | `pptx-viewer-core`           | PPTX-Dateien analysieren, erstellen, bearbeiten, serialisieren und konvertieren. Framework-unabhängig.        |
| [**React**](/react/)             | `pptx-react-viewer`          | React-Viewer, -Editor und -Presenter mit Symbolleiste, Inspektor, Zusammenarbeit und Export.                  |
| [**Vue 3**](/vue/)               | `pptx-vue-viewer`            | Vue-3-Viewer/Editor auf derselben Engine aufgebaut, mit demselben Funktionsumfang.                            |
| [**Angular**](/angular/)         | `pptx-angular-viewer`        | Angular-Viewer/Editor auf derselben Engine aufgebaut, mit demselben Funktionsumfang.                          |
| [**Vanilla JS**](/vanilla/)      | `pptx-vanilla-viewer`        | Binding ohne Framework auf derselben Engine: reines DOM, eine Factory-Funktion, keine Framework-Abhängigkeit. |
| [**Svelte**](/svelte/)           | `pptx-svelte-viewer`         | Svelte-5-Komponente auf derselben Engine aufgebaut, mit demselben Funktionsumfang.                            |
| [**Tools / MCP**](/packages/mcp) | `pptx-viewer-mcp`            | 73 PPTX-Tool-Funktionen, ein MCP-Server für KI-Agenten und der Y.Doc-Zusammenarbeit-Codec.                    |
| **Installer**                    | `@christophervr/pptx-viewer` | Interaktive CLI, die das richtige Viewer-Paket in Ihr Projekt einbindet.                                      |

### Wie die Teile zusammenpassen

Alle fünf UI-Binding-Pakete bauen auf einer gemeinsamen Rendering-Schicht auf, die wiederum auf Core aufbaut:

```
pptx-react-viewer   ┐
pptx-vue-viewer     │
pptx-angular-viewer ├── ooxml-ui/pptx ── ooxml-core/pptx
pptx-vanilla-viewer │                               ├── emf-converter
pptx-svelte-viewer  ┘                               └── mtx-decompressor
```

- **`pptx-viewer-core`** besitzt alles rund um das Dateiformat: die Lade- und Speicherpipelines, das typisierte Datenmodell, die Designauflösung, die Geometrie-Engine, die Verschlüsselung und die Markdown-/SVG-Konverter. Es hat keine UI und läuft überall, wo JavaScript läuft.
- **`ooxml-ui/pptx`** veröffentlicht gemeinsames Rendering, Styling, Webkomponenten, Browser-Laden und Export-Ansichten aus `src/ui/src/pptx`. Jedes Binding importiert diese öffentliche API. DOM-freie Bearbeitungs- und Zusammenarbeitsoperationen liegen in `ooxml-core/pptx/editor`; das private Paket `pptx-viewer-shared` ist nur noch eine Kompatibilitätsfassade.
- **Die Bindings** sind dünne View-Schichten. Jedes übersetzt dieselben gemeinsamen Renderdaten in die Sprache seines Frameworks (JSX, SFC-Templates, Angular-Templates, Svelte-Runes oder reines DOM). Darum rendern die fünf Bindings identisch und bleiben funktionsgleich.
- **`emf-converter`** (EMF/WMF-Metadatei zu PNG) und **`mtx-decompressor`** (MicroType-Express-Einbettungsschriften) sind eigenständige npm-Abhängigkeiten von Core.
- **`pptx-viewer-mcp`** baut auf `pptx-viewer-core` auf und stellt Tool-Aufruf- und Zusammenarbeits-Schnittstellen für KI-Agenten bereit.

## Minimale Beispiele

::: code-group

```ts [Load, edit, save]
import { PptxHandler } from 'pptx-viewer-core';

const handler = new PptxHandler();
const data = await handler.load(arrayBuffer);

// data.slides is a typed, mutable model
data.slides[0].elements.filter((el) => el.type === 'text').forEach((el) => console.log(el.text));

const bytes = await handler.save(data.slides); // => Uint8Array (.pptx)
```

```ts [Create from scratch]
import { PptxHandler } from 'pptx-viewer-core';

const { handler, data, createSlide } = await PptxHandler.create({
	title: 'Q4 Report',
	theme: { colors: { accent1: '#FF6B6B' } },
});

data.slides.push(createSlide('Blank').addText('Hello', { fontSize: 36 }).build());

const bytes = await handler.save(data.slides);
```

:::

## Nächste Schritte

- [Installation](/de/guide/installation) - Pakete installieren und lokale Entwicklung einrichten.
- [Schnellstart](/de/guide/quick-start) - durchgängige Anwendungsbeispiele.
- [Architektur](/de/guide/architecture) - wie die Lade-/Speicherpipelines und das Schichtendesign funktionieren.
- [Das PptxData-Modell](/de/guide/data-model) - das typisierte Modell, das Sie bearbeiten.
- [Core-Paketübersicht](/core/) - die Engine zum Analysieren, Bearbeiten und Serialisieren.
- [React-Paketübersicht](/react/) - die Viewer/Editor-Komponente.
- [Einschränkungen](/de/guide/limitations) - wichtige Hinweise, die vor der Einführung gelesen werden sollten.

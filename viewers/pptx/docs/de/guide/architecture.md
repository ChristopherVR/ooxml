---
title: Architektur
description: Wie pptx-viewer aufgebaut ist - die Lade- und Speicherpipelines, die aus Mixins zusammengesetzte Laufzeit, die Designauflösung, die Geometrie-Engine und die gemeinsame Rendering-Schicht, die jedes Framework-Binding speist.
---

# Architektur

`pptx-viewer` ist ein geschichtetes System. Framework-Bindings (React, Vue 3, Angular, Svelte und ein Vanilla-JS-Binding mit reinem DOM) übernehmen die UI; eine gemeinsame Rendering-Schicht liefert framework-unabhängige Logik; und die Core-Engine kümmert sich um alles rund um das Analysieren, Bearbeiten und Speichern von PowerPoint-Dateien. Jede Schicht hängt nur von der darunterliegenden ab.

## Überblick

```
+---------------------------------------------------------------+
|                      Framework bindings                       |
|   React    |   Vue 3   |  Angular  |  Svelte 5  |  Vanilla JS |
| pptx-react | pptx-vue  | pptx-ang. | pptx-svelte| pptx-vanilla|
+------------------------------+--------------------------------+
                               |
                               v
+---------------------------------------------------------------+
|            Shared rendering layer (ooxml-ui/pptx)        |
|     geometry, styles, gradients, charts, connectors, text     |
+------------------------------+--------------------------------+
                               |
                               v
+---------------------------------------------------------------+
|                Core engine (pptx-viewer-core)                 |
|  PptxHandler (public API)                                     |
|    -> Runtime (parsing, serialization, theme resolution)      |
|         -> Types, Geometry, Colour, Builders, Converter       |
+---------------------------------------------------------------+
```

## Framework-Bindings

Die Binding-Pakete sind dünne Darstellungsschichten. Sie verwenden vorberechnete Renderdaten aus `ooxml-ui/pptx` und übersetzen sie in frameworkspezifische Templates (JSX, Vue-SFCs, Angular-Komponenten, Svelte-5-Runes oder reine DOM-Aufrufe beim Vanilla-Binding). Folien werden als skaliertes HTML/SVG mit CSS-Transformationen gerendert, was scharfen Text bei jedem Zoom, native Barrierefreiheit und volle DOM-Interaktivität ergibt.

Jedes Binding stellt einen Viewer/Editor-Einstiegspunkt bereit, der Zustand, Bearbeitung, Laden, Export und Präsentationsmodus in der Sprache seines Frameworks orchestriert: Hooks in React, Composables in Vue, Services in Angular, Runes in Svelte und eine einfache Factory-Funktion mit imperativer Instanz-API in Vanilla JS. Weil alle dieselbe gemeinsame Schicht nutzen, ist die Rendering-Ausgabe in allen fünf identisch.

## Gemeinsame Rendering-Schicht (`ooxml-ui/pptx`)

Das gemeinsame Rendering liegt in `src/ui/src/pptx/` im Repository-Stamm und wird über `ooxml-ui/pptx` veröffentlicht. DOM-freie Dokumentoperationen liegen in `src/core/pptx/editor/`, und Tool-Schemas in `src/core/pptx/automation/schemas/`. Der Renderer umfasst:

- **Verbinder-Routing**: ein A\*-Router über einem Hindernisgraphen (`connector-router-astar.ts`, `connector-router-graph.ts`, jetzt im gemeinsamen Bereich `ooxml-core/geometry`) plus Pfadaufbau und Neuführung.
- **Diagrammmathematik**: Achsenbereiche, Kategoriepositionierung, kartesische/polare Plot-Builder, Box-Whisker-Statistik, Kombi-/Kursdiagramm-Zusammensetzung (die Modulfamilie `chart-*`).
- **Animations- und Morph-Engines**: Zeitleistenaufbau und Wiedergabe (`animation-timeline-*.ts`) sowie Morph-Übergangszuordnung und Geometrieinterpolation (`morph-matching.ts`, `morph-geometry-interp.ts`, `morph-text.ts`).
- **Text- und Stilauflösung**: Aufzählungsnummerierung, Füll-/Verlaufsstile, Bildeffektfilter, Kinsoku-Zeilenumbruchstile.
- **Mathematik**: Konverter von OMML zu MathML und von LaTeX zu OMML (`omml-to-mathml.ts`, `latex-to-omml.ts`, jetzt im gemeinsamen Bereich `ooxml-core/math`).
- **Editor-Verhalten**: Verlauf, Zwischenablage, Ausrichtungshilfen, Formatübertragung, Suchen und Ersetzen, Zusammenarbeits-Sync und Anwesenheit.

Benachbarte Verzeichnisse decken `export/`, `i18n/`, `loader/`, `theme/` und den optionalen Renderer `smartart-3d/` ab.

::: info Gemeinsame öffentliche Pakete
Jedes Binding deklariert `ooxml-ui` als Laufzeitabhängigkeit und importiert dessen PowerPoint-Einstiegspunkte. Angular nutzt dieselbe öffentliche API, ohne eine Quellkopie mitzuführen. Das private Paket `pptx-viewer-shared` bleibt eine Kompatibilitätsfassade und enthält keine Implementierung.
:::

## Core-Engine (`pptx-viewer-core`)

Das Core-Paket ist vollständig framework-unabhängig. Es läuft in jeder JavaScript-Umgebung: Browser, Node.js, Web Worker oder Serverless-Funktion. Sein öffentlicher Einstiegspunkt ist `PptxHandler`.

> **Wo der Code liegt.** Die Engine ist der Bereich `pptx` des gemeinsamen, öffentlichen Pakets `ooxml-core` (`src/core/pptx/`, Subpfade `/pptx`, `/pptx/converter`, `/pptx/cli`, `/pptx/signature-node`); der Viewer-Ordner enthält die Framework-Adapter. `pptx-viewer-core` ist ein dünnes Paket, das diese Subpfade mit unveränderter öffentlicher API erneut exportiert; daher sind alle folgenden Pfade, die mit `src/core/pptx/` beginnen, relativ zum Monorepo-Stamm. Die Engine verwendet weiterhin ein eigenes `fast-xml-parser`-Objektmodell und wird mit gelockerten TypeScript-Optionen kompiliert; die Vereinheitlichung ihres XML-Modells mit dem gemeinsamen Bereich `xml` und das Verschärfen der Optionen sind die nächsten Schritte der Migration. Farbprimitive, Voreinstellungs-Formgeometrie und EMU-Konstanten werden mit den anderen Office-Bereichen geteilt (`color`, `geometry`, `units`).

### Die Fassade und die aus Mixins zusammengesetzte Laufzeit

Die Engine besteht aus drei Schichten mit abnehmender Oberfläche:

```
PptxHandler                    static factories (create / createBlank)
  └─ PptxHandlerCore           thin facade: load / save / export / encryption
       └─ IPptxHandlerRuntime  the actual engine, assembled from many mixin modules
```

- **`PptxHandler`** (`src/core/pptx/core/PptxHandler.ts`) ergänzt die statischen Builder-Einstiegspunkte `create()` / `createBlank()`.
- **`PptxHandlerCore`** delegiert das gesamte aufwendige Analysieren, Serialisieren und die XML-Manipulation an eine injizierte `IPptxHandlerRuntime`. Die Laufzeit ist über Konstruktorabhängigkeiten (`runtime` oder `runtimeFactory`) austauschbar; so tauschen Tests und alternative Hosts Implementierungen aus.
- **`PptxHandlerRuntime`** ist nicht eine Klasse in einer Datei. Sie ist aus vielen **fokussierten Modulen** in `src/core/pptx/core/core/runtime/` zusammengesetzt, jeweils benannt `PptxHandlerRuntime<Concern>.ts` und jeweils für genau ein Anliegen zuständig: `PptxHandlerRuntimeChartParsing.ts`, `PptxHandlerRuntimeThemeLoading.ts`, `PptxHandlerRuntimeSaveElementWriter.ts`, `PptxHandlerRuntimeSmartArtParsing.ts` und so weiter.

Jedes Modul deklariert eine Klasse, die die vom vorherigen Modul exportierte Klasse erweitert. So entsteht eine lineare Vererbungskette, die Fähigkeit auf Fähigkeit schichtet:

```ts
// PptxHandlerRuntimeLoadPipeline.ts
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeLoadSession';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	// adds the load-pipeline capability on top of everything below it
}
```

Neue Engine-Fähigkeiten (auch neue Elementtypen) werden als neue Glieder dieser Kette ergänzt, statt bestehende Dateien wachsen zu lassen. Übergreifende Mitspieler (Kompatibilitätswarnungen, XML-Factorys, Content-Type-Builder) sind hinter `I*`-Service-Schnittstellen definiert und über eine Abhängigkeits-Factory injiziert, sodass jedes Mixin isoliert testbar bleibt.

### Die Ladepipeline

`handler.load(arrayBuffer)` durchläuft diese Stufen:

```
ArrayBuffer
  │  container sniff: an OLE compound file is either a legacy binary
  │  .ppt (converted to an in-memory .pptx package) or an encrypted
  │  OOXML package (decrypted with the supplied password)
  │  ZIP signature check on whatever comes out
  ▼
JSZip.loadAsync                 in-memory archive
  │  zip-bomb guard: 500 MiB uncompressed budget (configurable),
  │  hard cap of 65,536 entries
  ▼
fast-xml-parser                 XML parts -> JS object trees
  │  Strict OOXML detection: strict namespace URIs are transparently
  │  normalized to Transitional for all subsequent parses
  ▼
Theme / master / layout resolution
  │  colour maps, font schemes, format schemes, placeholder styles
  ▼
Per-slide shape-tree (spTree) parsing
  │  each <p:sp>, <p:pic>, <p:graphicFrame>, ... becomes a typed PptxElement
  ▼
PptxLoadDataBuilder             assembles the final model
  ▼
PptxData
```

1. **Zuerst Container-Prüfungen.** Eine OLE-Verbunddatei ist nicht zwingend ein Fehler: So kommen sowohl ein altes binäres `.ppt` als auch ein verschlüsseltes OOXML-Paket an. Eine PowerPoint-97-2003-Präsentation (erkennbar am Stream `PowerPoint Document`) wird von `core/ppt/` in ein äquivalentes `.pptx`-Paket im Arbeitsspeicher umgewandelt und von dort geladen, sodass der Rest der Pipeline nie erfährt, dass die Präsentation binär war. Andernfalls gilt der Container als verschlüsselt und löst `EncryptedFileError` aus, sofern kein Passwort zur Entschlüsselung angegeben ist; ein passwortgeschütztes `.ppt` löst `EncryptedPptError` aus, weil die alte RC4-Verschlüsselung nicht entschlüsselt wird. Eingaben, die weder ZIP noch OLE-Verbunddatei sind, werden mit einer klaren Fehlermeldung abgelehnt. Überdimensionierte Archive lösen `ZipBombError` aus, bevor irgendein Parsing stattfindet.
2. **XML-Parsing** verwendet `fast-xml-parser`; jeder Teil wird zu einem einfachen Objektbaum (Typ `XmlObject`) mit Attributen mit Präfix `@_`.
3. **Designauflösung** lädt je Master das Design, die Farbzuordnung (`p:clrMap`), das Schriftschema und das Formatschema, damit das Folien-Parsing Schemafarben und Stilverweise auflösen kann (siehe unten).
4. **Element-Parsing** macht aus dem Formenbaum jeder Folie die [diskriminierte Union `PptxElement`](/de/guide/data-model#die-pptxelement-union), rechnet EMU-Koordinaten in Pixel um und bewahrt rohes XML für Konstrukte, die das typisierte Modell nicht abdeckt.
5. **`PptxLoadDataBuilder`** stellt alles Übrige zusammen: Abschnitte, benutzerdefinierte Bildschirmpräsentationen, eingebettete Schriften, Notizen- und Handzettelmaster, Tags, Kommentarautoren, Dokumenteigenschaften, das Vorschaubild und Kompatibilitätswarnungen.

### Die Speicherpipeline

`handler.save(slides, options?)` kehrt den Vorgang um (`PptxHandlerRuntimeSavePipeline.ts`):

```
PptxSlide[]
  │  resolve conformance class ('preserve' | 'strict' | 'transitional')
  ▼
Reconcile presentation slide list      order, additions, deletions, rels
  ▼
Serialize each slide to OpenXML        elements -> <p:spTree>, embed new media
  ▼
Rebuild [Content_Types].xml            slide overrides + media defaults
  ▼
Comments, masters, layouts             typed mutations applied; untouched
  │                                    parts pass through verbatim
  ▼
Optional Strict conversion             remap namespaces if target is strict
  ▼
JSZip -> Uint8Array                    a valid .pptx / .ppsx / .pptm
```

Zwei Eigenschaften dieser Pipeline sind für die Treue wichtig:

- **Standardmäßig Durchreichen.** Teile, die Sie nicht bearbeitet haben (Master, Layouts, Notizen, unbekannte Erweiterungen, herstellerspezifisches Markup), werden byteweise aus dem Originalarchiv übernommen oder aus ihren erhaltenen Parsebäumen neu ausgegeben. Speichern ist ein gezieltes Umschreiben, keine vollständige Neuerzeugung.
- **Konformitätsbewusste Ausgabe.** Das gespeicherte Paket entspricht standardmäßig der OOXML-Konformitätsklasse der geladenen Datei und kann auf Strict oder Transitional erzwungen werden. Was genau umgeschrieben wird, steht unter [OpenXML-Konformität](/architecture/openxml-conformance).

### Auflösungskette für Designs

PowerPoint-Stile werden über eine Vererbungskette aufgelöst, und die Engine bildet sie exakt nach:

```
Element  ->  Placeholder  ->  Layout  ->  Master  ->  Theme
```

- Explizite Eigenschaften eines Elements gewinnen immer.
- Ein Platzhalterelement (Titel, Textkörper, Fußzeile, ...) erbt Text- und Positionsvorgaben vom passenden Platzhalter in seinem **Layout**, das wiederum vom **Master** erbt (`p:txStyles` für Standardwerte von Titel-/Körper-/sonstigem Text).
- Stilverweise (`a:fillRef`, `a:lnRef`, `a:effectRef`, `a:fontRef`) werden per Index in das Formatschema des **Designs** aufgelöst, wobei die Farbe des Verweises die Platzhalterfarbe des Designs ersetzt.
- Schemafarben (`accent1`, `bg1`, `tx1`, ...) werden über das `p:clrMap` des Masters (und ein etwaiges `p:clrMapOvr` des Layouts) in das Farbschema des Designs aufgelöst.

Die Kette wird durch eigene Laufzeit-Mixins umgesetzt: `PptxHandlerRuntimeThemeLoading`, `...ThemeProcessing`, `...ThemeFormatScheme`, `...ThemeRefResolution`, `...ThemeOverrides` und das Platzhalter-Trio `...PlaceholderLookup`, `...PlaceholderDefaults`, `...PlaceholderStyles`. Mehrere Master werden unterstützt, jeder mit eigener Farbzuordnung und eigenem Formatschema.

### Geometrie-Engine

`src/core/pptx/core/geometry/` (plus die gemeinsamen Voreinstellungstabellen in `src/core/geometry/`) macht aus DrawingML-Geometrie renderbare Pfade:

- **Voreinstellungsformen**: Definitionen für den Voreinstellungsformen-Katalog von ECMA-376, nach Familien gruppiert (`preset-shape-definitions-arrows.ts`, `-flowchart.ts`, `-action-buttons.ts`, `-callouts`, ...), jeweils mit den Hilfsformeln der Spezifikation ausgedrückt.
- **Auswertung der Hilfsformeln** (`guide-formula-eval.ts` und Umgebung): setzt die Formelsprache von ECMA-376 um (`*/`, `+-`, `pin`, `at2`, `cos`, ...), damit die Formgeometrie korrekt auf Anpassungswerte reagiert, die gelben Rautengriffe, die Sie in PowerPoint ziehen können.
- **Benutzerdefinierte Geometrie** (`custom-geometry-parser.ts`, `freeform-builder.ts`): analysiert `a:custGeom`-Pfadbefehle und baut Freiformen.
- **Clip-Pfade** (`preset-clip-paths-core.ts`, `-extended.ts`): erzeugen CSS/SVG-Clip-Pfade, damit HTML-Inhalte (Bilder, Text) auf jede Voreinstellungsform zugeschnitten werden können.
- **Verbindergeometrie** (`connector-geometry.ts`): gewinkelte und gebogene Verbinderformen, Pfeilspitzen und Spiegelungsbehandlung. (Das **Routing** von Live-Verbindern um Hindernisse liegt im A\*-Router der gemeinsamen Schicht.)

### Konverter

`src/core/pptx/converter/` setzt die PPTX-zu-Markdown-Konvertierung mit einem Registry-Muster um: Jeder Elementtyp hat einen Prozessor (`shape-element-processor`, `table-element-processor`, `ole-element-processor`, ...), registriert für seine `type`-Diskriminante, und `PptxMarkdownConverter` verteilt pro Element. Im selben Verzeichnis liegen der SVG-Exporter und der OMML-zu-LaTeX-Konverter für Formeln.

## Zentrale Entwurfsentscheidungen

| Entscheidung                              | Begründung                                                                                                                                                                                                                                           |
| ----------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **CSS-basiertes Rendering** (kein Canvas) | Scharfer Text bei jedem Zoom, native Barrierefreiheit, DOM-Interaktivität und Standard-CSS-Styling.                                                                                                                                                  |
| **Aus Mixins zusammengesetzte Engine**    | Viele kleine Laufzeitmodule mit je einem Anliegen halten jede Fähigkeit isoliert und testbar. Neue Fähigkeiten (auch neue Elementtypen) kommen als neue Mixins hinzu, nicht als größere Dateien.                                                     |
| **Diskriminierte Union für Elemente**     | TypeScript grenzt über das Feld `type` auf den richtigen Elementtyp ein und liefert volle Typsicherheit ohne Casting.                                                                                                                                |
| **Auflösungskette für Designs**           | Element, Platzhalter, Layout, Master, Design bildet die Stilvererbung von PowerPoint selbst nach.                                                                                                                                                    |
| **EMU als interne Einheit**               | PowerPoint verwendet English Metric Units (914.400 EMU pro Zoll; 9.525 EMU pro Pixel bei 96 DPI). Analysierte Elemente stellen Pixelwerte für bequeme Layoutrechnung bereit; exakte EMU-Werte bleiben erhalten, wo es die Roundtrip-Treue erfordert. |
| **Durchreich-Speichern**                  | Nur bearbeitete Teile werden neu geschrieben; alles andere wird unverändert übernommen, sodass unbekanntes Markup und Herstellererweiterungen erhalten bleiben.                                                                                      |
| **Gemeinsame Logik, dünne Bindings**      | Die gesamte framework-unabhängige Viewer-Logik liegt einmal in `ooxml-ui/pptx`; Bindings sind reine View-Schichten, was die fünf Frameworks auf demselben Stand hält.                                                                                |

## Weiterführend

- [Das PptxData-Modell](/de/guide/data-model) - die vollständige Form einer analysierten Präsentation, Elementtypen und Einheiten.
- [OpenXML-Konformität](/architecture/openxml-conformance) - Behandlung von Strict und Transitional und der Konformitätsvertrag.
- [Core-Paketübersicht](/core/) - die öffentliche API-Referenz.

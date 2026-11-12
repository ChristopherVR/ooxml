---
title: Viewport-Anpassung
description: Die Anpassung der Folien in einem eigenen Host konfigurieren, ohne die festgelegten Foliengrößen oder den Benutzerzoom zu verändern.
---

# Viewport-Anpassung

Alle fünf UI-Bindings akzeptieren die optionalen Host-Optionen `fitPadding` und `maxFitScale`. Verwenden Sie `fitPadding: 0` und `maxFitScale: null`, wenn der Host bereits eigene Abstände bereitstellt und die Folie über ihre ursprüngliche Pixelgröße hinauswachsen darf. Werden diese Optionen weggelassen, bleibt die bisherige Anpassungsregel erhalten.

## Optionen und Standardwerte {#options-and-defaults}

`fitPadding` ist ein nicht skalierter Abstand in CSS-Pixeln **auf jeder Seite**: `8` reserviert 16 Pixel in jeder Achse, `{ horizontal: 4, vertical: 16 }` reserviert 8 Pixel horizontal und 32 Pixel vertikal. `maxFitScale` ist eine positive Obergrenze des Anpassungsfaktors; `null` bedeutet unbegrenzt. Ungültige Werte fallen auf die gemeinsamen Standardwerte zurück.

Alle fünf Bindings verwenden dieselben Standardwerte, nämlich die des React-Editors: Ein Desktop-Viewport (768 px oder breiter) reserviert 4 / 16 px pro Seite und begrenzt den Anpassungsfaktor auf 1. Ein Viewport, der schmaler als 768 px ist, reserviert stattdessen 12 / 12 px pro Seite.

| Bindung | Abstand pro Seite horizontal / vertikal | Maximaler Anpassungsfaktor |
| ------- | --------------------------------------- | -------------------------- |
| React   | 4 / 16 px (12 / 12 px unter 768 px)     | 1                          |
| Vue     | 4 / 16 px (12 / 12 px unter 768 px)     | 1                          |
| Angular | 4 / 16 px (12 / 12 px unter 768 px)     | 1                          |
| Svelte  | 4 / 16 px (12 / 12 px unter 768 px)     | 1                          |
| Vanilla | 4 / 16 px (12 / 12 px unter 768 px)     | 1                          |

Jedes UI-Paket exportiert die Typen `ViewportFitOptions` und `ViewportFitPadding`. Diese Optionen betreffen die normale Anpassung im Viewer, nicht die festgelegten Foliengrößen, gespeicherte Inhalte, Miniaturansichten, den Export, den Präsentationsmodus oder die separate Benutzer-Zoomeinstellung. Vanilla schreibt den aufgelösten Abstand immer inline auf seinen Viewport, daher hat eine Stylesheet-Regel für das Padding dieses Elements keine Wirkung; verwenden Sie stattdessen `fitPadding`.

## Messung {#measurement}

Jede Bindung passt die Folie an die fraktionale Clientgröße des Viewports an (seinen Padding-Bereich ohne Scrollbalken, mit der Genauigkeit, mit der der Browser ihn layoutet), nicht an die ganzzahligen Werte `clientWidth` und `clientHeight`. Ein Viewport, der mit 801,5 px layoutet wird, wird auf 801,5 px angepasst. Mit `fitPadding: 0` füllt die Folie ihn daher genau aus und ragt nie darüber hinaus. Rahmen werden mit ihrer tatsächlichen Breite abgezogen, wodurch ihr Bruchteil bei einem Browser-Zoom erhalten bleibt (ein `1px`-Rahmen ist bei 90 % 1,111 px breit).

Die exakte Größe wird nur in einem einfachen Kontext verwendet, in dem nichts den Viewport transformiert oder per CSS zoomt (Browser-Zoom ist unproblematisch). Dort entspricht das Begrenzungsrechteck des Viewports seiner berechneten `width` und `height`, die in seinen eigenen CSS-Pixeln angegeben sind. Wenn eine CSS-`transform`- oder `zoom`-Eigenschaft am Viewport oder an einem Vorfahren das Rechteck skaliert, wird die Folie ebenfalls durch diese Skalierung layoutet und gezeichnet, und ihre eigene Rundung im Layout kann ihre Kanten um einen Bruchteil eines Pixels verschieben. Die Anpassung reserviert dann einen Pixel, damit die Folie innerhalb des Viewports bleibt und der Viewport nicht scrollt. Ein Pixel wird auch reserviert, solange der Viewport eine Scrollbar anzeigt, da die exakte Größe dann nicht zuverlässig ermittelt werden kann.

Spätere Änderungen erfasst ein `ResizeObserver` sowohl für den Content-Box- als auch für den Border-Box-Bereich des Viewports. Daher passt sich die Folie auch an, wenn nur das Padding des Viewports geändert wird.

## Viewer konfigurieren {#configure-your-viewer}

Geben Sie dem Host eine echte Breite und Höhe. Ein Flex-Element benötigt unter Umständen zusätzlich `min-height: 0`, damit es in den verfügbaren Platz passt. Ein Fit-Padding von null gibt keinen Platz frei, der von Werkzeugleisten, Seitenleisten oder anderen Host-Elementen belegt wird.

### React {#react}

```tsx
<PowerPointViewer content={content} fitPadding={0} maxFitScale={null} />
```

Dieselben Optionen können an `useViewerBuildingBlocks` übergeben werden, wenn `Toolbar` und `SlideCanvas` in eine eigene Oberfläche eingebunden werden. Beispiel:

```tsx
const { canvasProps, toolbarProps } = useViewerBuildingBlocks({
	content,
	canEdit: true,
	fitPadding: 0,
	maxFitScale: null,
});
```

Rendern Sie die zurückgegebenen Props wie üblich. Die Anpassung nutzt den gemessenen Canvas-Viewport, nicht den gesamten Host einschließlich der Werkzeugleiste. Bestehende Lineal-Versätze bleiben wirksam. Setzen Sie `showRulers={false}` auf einem eigenen `SlideCanvas`, wenn er genau an den Viewport-Kanten ausgerichtet sein muss.

`SlideCanvas` verfolgt seinen Viewport automatisch, einschließlich bedingtem Entfernen und erneutem Einhängen, während der Building-Block-Hook eingehängt bleibt. Wenn Sie `useZoomViewport` aus `pptx-react-viewer/internals` mit einem eigenen Viewport-Element verwenden, binden Sie `ref={zoom.setCanvasViewportNode ?? zoom.canvasViewportRef}`, damit Anpassung und Zoom per Mausrad dem aktuellen Element folgen. Das bestehende Objekt `canvasViewportRef` bleibt zum Lesen des Viewports verfügbar.

### Vue {#vue}

```vue
<PowerPointViewer :content="content" :fit-padding="0" :max-fit-scale="null" />
```

Dieselben Props stehen auf `SlideCanvas` zur Verfügung.

### Angular {#angular}

```html
<pptx-viewer [content]="content()" [fitPadding]="0" [maxFitScale]="null" />
```

### Svelte {#svelte}

```svelte
<PowerPointViewer {source} fitPadding={0} maxFitScale={null} />
```

### Vanilla JS {#vanilla-js}

```ts
const viewer = createPptxViewer(container, {
	source,
	fitPadding: 0,
	maxFitScale: null,
});
```

Das bestehende Lineal-Layout bleibt in jedem Binding erhalten. Deaktivieren Sie die Lineale, wenn der Host eine an den Kanten ausgerichtete Anpassung benötigt. Eine Änderung von `fitPadding` entfernt weder den Platz der Lineale noch ändert sie die übrigen Layout-Einstellungen des Viewers.

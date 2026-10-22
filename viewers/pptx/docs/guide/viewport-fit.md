---
title: Viewport Fitting
description: Configure slide fitting inside a custom host without changing authored slide dimensions or user zoom.
---

# Viewport fitting

All five UI bindings accept optional `fitPadding` and `maxFitScale` host options.
Use `fitPadding: 0` and `maxFitScale: null` when the host already supplies its own
gutters and the slide should be allowed to grow beyond its original pixel size.
Omitting these options preserves the existing fit policy.

## Options and defaults

`fitPadding` is an unscaled CSS-pixel allowance **on each side**: `8` reserves
16 pixels in each axis; `{ horizontal: 4, vertical: 16 }` reserves 8 pixels
horizontally and 32 vertically. `maxFitScale` is a positive fit-factor ceiling;
`null` means unlimited. Invalid values fall back to the binding's defaults.

| Binding | Per-side horizontal / vertical padding | Maximum fit scale |
| ------- | -------------------------------------- | ----------------- |
| React   | 4 / 16 px                              | 1                 |
| Vue     | 8 / 16 px                              | 1                 |
| Angular | 8 / 16 px                              | 1                 |
| Svelte  | 24 / 24 px                             | Unlimited         |
| Vanilla | 16 / 16 px                             | Unlimited         |

Each UI package exports the `ViewportFitOptions` and `ViewportFitPadding` types.
These options affect ordinary viewer fitting, not authored slide dimensions,
saved content, thumbnails, export, presentation mode, or the separate user zoom
setting. Vanilla preserves its existing mobile CSS padding when `fitPadding`
is omitted.

## Measurement

Every binding fits to the viewport's fractional client size (its padding box
without scrollbars, at the precision the browser lays it out), not the
whole-pixel `clientWidth` and `clientHeight`. A viewport laid out at 801.5px is
fitted to 801.5px, so with `fitPadding: 0` the slide fills it exactly and never
extends past it. Borders are taken out at their used width, which keeps its
fraction under browser zoom (a `1px` border is 1.111px at 90%).

The exact size is used only in a plain context, where nothing transforms or
CSS-zooms the viewport (browser zoom is fine). There the viewport's bounding
rect matches its computed `width` and `height`, which are in its own CSS
pixels. When a CSS `transform` or `zoom` on the viewport or an ancestor scales
the rect, the slide is also laid out and painted through that scale, and its own
layout rounding can move its edges by a fraction of a pixel. The fit then
reserves one pixel instead, so the slide stays inside the viewport and the
viewport does not scroll. One pixel is also reserved while the viewport shows a
scrollbar, since the exact size cannot be recovered then.

Later changes are picked up by `ResizeObserver`s on both the content box and the
border box of the viewport, so a change to the viewport's padding alone also
refits the slide.

## Configure your viewer

Give the host a real width and height. A flex child may also need `min-height: 0`
so it can fit within the available space. Zero fit padding cannot recover space
used by toolbars, side panels, or other host chrome.

### React

```tsx
<PowerPointViewer content={content} fitPadding={0} maxFitScale={null} />
```

The same options can be passed to `useViewerBuildingBlocks` when composing
`Toolbar` and `SlideCanvas` into a custom shell. For example:

```tsx
const { canvasProps, toolbarProps } = useViewerBuildingBlocks({
	content,
	canEdit: true,
	fitPadding: 0,
	maxFitScale: null,
});
```

Render the returned props as usual; fitting uses the measured canvas viewport,
not the entire host including its toolbar. Existing ruler offsets remain in
effect. Set `showRulers={false}` on a custom `SlideCanvas` when it must align
directly with the viewport edges.

`SlideCanvas` tracks its viewport automatically, including conditional unmounts
and remounts while the building-block hook stays mounted. If you use
`useZoomViewport` from `pptx-react-viewer/internals` with your own viewport element, bind
`ref={zoom.setCanvasViewportNode ?? zoom.canvasViewportRef}` so fitting and native
wheel zoom follow the current element. The existing `canvasViewportRef` object
remains available for reading the viewport.

### Vue

```vue
<PowerPointViewer :content="content" :fit-padding="0" :max-fit-scale="null" />
```

The same props are available on `SlideCanvas`.

### Angular

```html
<pptx-viewer [content]="content()" [fitPadding]="0" [maxFitScale]="null" />
```

### Svelte

```svelte
<PowerPointViewer {source} fitPadding={0} maxFitScale={null} />
```

### Vanilla JS

```ts
const viewer = createPptxViewer(container, {
	source,
	fitPadding: 0,
	maxFitScale: null,
});
```

Existing ruler layout is preserved in each binding. Disable rulers when the
host needs edge-aligned fitting; changing `fitPadding` does not remove ruler
space or change the viewer's other layout controls.

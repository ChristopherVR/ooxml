/**
 * Framework metadata + code samples for the landing page. None of this is
 * localized: package names, entry points and code are identical in every
 * locale. Samples mirror docs/frameworks/*.md and packages/bindings/README.md;
 * keep them in sync with those pages. `entry` is the npm package name of each
 * framework adapter.
 */

export interface FrameworkSample {
	id: string;
	label: string;
	/** npm package name of the adapter. */
	entry: string;
	file: string;
	docsHref: string;
	code: string;
}

export const FRAMEWORKS: FrameworkSample[] = [
	{
		id: 'react',
		label: 'React',
		entry: 'visio-react-viewer',
		file: 'Diagram.tsx',
		docsHref: '/frameworks/react',
		code: `import { VisioViewer } from 'visio-react-viewer';
import type { VisioDocument } from 'visio-react-viewer';

export function Diagram({ diagram }: {
  diagram: VisioDocument;
}) {
  return <VisioViewer
    document={diagram}
    style={{ height: 600 }}
    events={{ 'document-error': console.error }}
  />;
}
// The native component handles mount and cleanup.`,
	},
	{
		id: 'vue',
		label: 'Vue 3',
		entry: 'visio-vue-viewer',
		file: 'Diagram.vue',
		docsHref: '/frameworks/vue',
		code: `<script setup lang="ts">
import { VisioViewer } from 'visio-vue-viewer';
import type { VisioDocument } from 'visio-vue-viewer';
defineProps<{ diagram: VisioDocument }>();
</script>

<template>
  <VisioViewer
    :document="diagram"
    style="height: 600px"
    @document-error="console.error"
  />
</template>`,
	},
	{
		id: 'angular',
		label: 'Angular',
		entry: 'visio-angular-viewer',
		file: 'diagram.ts',
		docsHref: '/frameworks/angular',
		code: `import { Component, Input } from '@angular/core';
import { VisioViewerComponent }
  from 'visio-angular-viewer';
import type { VisioDocument } from 'visio-angular-viewer';

@Component({
  selector: 'app-diagram',
  imports: [VisioViewerComponent],
  template: '<visio-viewer-host [document]="diagram" />',
  styles: ['visio-viewer-host { display:block; height:600px }'],
})
export class Diagram {
  @Input() diagram: VisioDocument | null = null;
}`,
	},
	{
		id: 'svelte',
		label: 'Svelte',
		entry: 'visio-svelte-viewer',
		file: 'Diagram.svelte',
		docsHref: '/frameworks/svelte',
		code: `<script lang="ts">
import VisioViewer
  from 'visio-svelte-viewer';
import type { VisioDocument } from 'visio-svelte-viewer';
let { diagram }: { diagram: VisioDocument } = $props();
</script>

<VisioViewer
  document={diagram}
  style="height:600px"
  events={{ 'document-error': console.error }}
/>`,
	},
	{
		id: 'solid',
		label: 'Solid',
		entry: 'visio-solid-viewer',
		file: 'Diagram.tsx',
		docsHref: '/frameworks/solid',
		code: `import { VisioViewer } from 'visio-solid-viewer';
import type { VisioDocument } from 'visio-solid-viewer';

export function Diagram(props: { diagram: VisioDocument }) {
  return <VisioViewer
    document={props.diagram}
    style={{ height: '600px' }}
    events={{ 'document-error': console.error }}
  />;
}
// Each binding uses the same shared browser element.`,
	},
	{
		id: 'vanilla',
		label: 'Vanilla JS',
		entry: 'visio-vanilla-viewer',
		file: 'main.ts',
		docsHref: '/frameworks/vanilla',
		code: `import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
  events: { 'document-error': console.error },
});

await viewer.load(file); // A local File or bytes
viewer.fit();

// When your view unmounts:
viewer.destroy();`,
	},
];

/** Export the current page and an explicit VSDX copy through the shared handle. */
export const CORE_SAMPLE = `import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host);
await viewer.load(file);

// Current page as a static SVG: string, dimensions, byteLength, diagnostics.
const { svg, diagnostics } = viewer.exportSvg();

// Original or edited VSDX bytes for an explicit download; never
// overwrites the source file.
const { bytes, dirty } = await viewer.exportVsdx();`;

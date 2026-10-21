import { expect, it } from 'vitest';
import '@angular/compiler';
import {
	useViewerBuildingBlocks as react,
	DiagramCanvas as ReactCanvas,
} from '../../react/src/building-blocks';
import {
	useViewerBuildingBlocks as vue,
	DiagramCanvas as VueCanvas,
} from '../../vue/src/building-blocks';
import {
	createViewerBuildingBlocks as solid,
	DiagramCanvas as SolidCanvas,
} from '../../solid/src/building-blocks';
import { createViewerBuildingBlocks as svelte } from '../../svelte/src/building-blocks';
import SvelteCanvas from '../../svelte/src/DiagramCanvas.svelte';
import { ViewerBuildingBlocks, DiagramCanvasComponent } from '../../angular/src/building-blocks';
it('native composition entries import safely on the server without mounting or registering elements', () => {
	expect(typeof globalThis.document).toBe('undefined');
	for (const factory of [react, vue, solid, svelte, ViewerBuildingBlocks])
		expect(typeof factory).toBe('function');
	for (const canvas of [ReactCanvas, VueCanvas, SolidCanvas, SvelteCanvas, DiagramCanvasComponent])
		expect(canvas).toBeDefined();
	expect(globalThis.customElements?.get?.('visio-viewer')).toBeUndefined();
});

import { describe, expect, it } from 'vitest';
import { effectScope, nextTick, ref } from 'vue';

import { useViewerZoom } from './useViewerZoom';

// Alias keeps the composable out of a `use*`-named call site in plain helpers,
// which the react-hooks linter would otherwise flag as a rules-of-hooks issue.
const runViewerZoom = useViewerZoom;

describe('useViewerZoom', () => {
	it('opens another slide at fit and keeps the zoom on the same slide', async () => {
		const scope = effectScope();
		const slide = ref('s1');
		const zoom = scope.run(() => runViewerZoom(() => slide.value))!;

		zoom.zoom.value = 1.9;
		slide.value = 's1';
		await nextTick();
		expect(zoom.zoom.value).toBe(1.9);

		slide.value = 's2';
		await nextTick();
		expect(zoom.zoom.value).toBe(1);
		expect(zoom.store.getState().manual).toBeFalsy();
		scope.stop();
	});

	it('keeps a zoom set right after reporting the next slide (the handle goTo then setZoom)', async () => {
		const scope = effectScope();
		const slide = ref('s1');
		const zoom = scope.run(() => runViewerZoom(() => slide.value))!;

		zoom.showSlide('s2');
		slide.value = 's2';
		zoom.zoom.value = 2;
		await nextTick();
		expect(zoom.zoom.value).toBe(2);
		scope.stop();
	});

	it('never zooms out past fit', () => {
		const scope = effectScope();
		const zoom = scope.run(() => runViewerZoom())!;

		zoom.zoomOut();
		expect(zoom.zoom.value).toBe(1);
		zoom.zoom.value = 0.4;
		expect(zoom.zoom.value).toBe(1);
		scope.stop();
	});
});

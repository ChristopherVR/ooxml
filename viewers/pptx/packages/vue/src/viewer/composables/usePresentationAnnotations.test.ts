// oxlint-disable react-hooks/rules-of-hooks
import { describe, expect, it } from 'vitest';
import { ref } from 'vue';

import { usePresentationAnnotations } from './usePresentationAnnotations';

// The stroke, eraser and per-slide rules are the shared
// `createPresentationInkController`'s (tested in `ooxml-ui/pptx`); these pin
// the Vue adapter's wiring onto it.
describe('usePresentationAnnotations', () => {
	it('toggles a tool on/off and ignores tool changes while inactive', () => {
		const active = ref(true);
		const a = usePresentationAnnotations({ isActive: active, activeSlideIndex: ref(0) });
		a.setPresentationTool('pen');
		expect(a.presentationTool.value).toBe('pen');
		// Selecting the same tool again returns to none.
		a.setPresentationTool('pen');
		expect(a.presentationTool.value).toBe('none');

		active.value = false;
		a.setPresentationTool('pen');
		expect(a.presentationTool.value).toBe('none');
	});

	it('draws a pen stroke and commits it on pointer up', () => {
		const a = usePresentationAnnotations({ isActive: ref(true), activeSlideIndex: ref(0) });
		a.setPresentationTool('pen');
		a.handlePointerDown(0, 0);
		a.handlePointerMove(10, 10);
		a.handlePointerMove(20, 20);
		expect(a.currentStroke.value?.points).toHaveLength(3);
		a.handlePointerUp();
		expect(a.currentStroke.value).toBeNull();
		expect(a.annotationStrokes.value).toHaveLength(1);
		expect(a.hasAnyAnnotations.value).toBeTruthy();
	});

	it('discards single-point strokes', () => {
		const a = usePresentationAnnotations({ isActive: ref(true), activeSlideIndex: ref(0) });
		a.setPresentationTool('pen');
		a.handlePointerDown(5, 5);
		a.handlePointerUp();
		expect(a.annotationStrokes.value).toHaveLength(0);
	});

	it('does not draw while inactive', () => {
		const active = ref(false);
		const a = usePresentationAnnotations({ isActive: active, activeSlideIndex: ref(0) });
		// Tool stays none while inactive, so no stroke begins.
		a.setPresentationTool('pen');
		a.handlePointerDown(0, 0);
		expect(a.currentStroke.value).toBeNull();
	});

	it('tracks the laser position only with the laser tool active', () => {
		const a = usePresentationAnnotations({ isActive: ref(true), activeSlideIndex: ref(0) });
		a.handleLaserMove(1, 1);
		expect(a.laserPosition.value).toBeNull();
		a.setPresentationTool('laser');
		a.handleLaserMove(42, 24);
		expect(a.laserPosition.value).toStrictEqual({ x: 42, y: 24 });
		a.handleLaserLeave();
		expect(a.laserPosition.value).toBeNull();
	});

	it('erases strokes near a point with the eraser tool', () => {
		const a = usePresentationAnnotations({ isActive: ref(true), activeSlideIndex: ref(0) });
		a.setPresentationTool('pen');
		a.handlePointerDown(0, 0);
		a.handlePointerMove(1, 1);
		a.handlePointerUp();
		expect(a.annotationStrokes.value).toHaveLength(1);

		a.setPresentationTool('eraser');
		a.eraseAtPoint(0, 0);
		expect(a.annotationStrokes.value).toHaveLength(0);
		expect(a.hasAnyAnnotations.value).toBeFalsy();
	});

	it('saves & restores strokes per slide', () => {
		const idx = ref(0);
		const a = usePresentationAnnotations({ isActive: ref(true), activeSlideIndex: idx });
		a.setPresentationTool('pen');
		a.handlePointerDown(0, 0);
		a.handlePointerMove(1, 1);
		a.handlePointerUp();
		expect(a.annotationStrokes.value).toHaveLength(1);

		// Move to slide 1: its strokes are empty.
		idx.value = 1;
		expect(a.annotationStrokes.value).toHaveLength(0);

		// Move back: slide 0's stroke is restored.
		idx.value = 0;
		expect(a.annotationStrokes.value).toHaveLength(1);

		// All-slide map carries the slide-0 stroke.
		expect(a.allSlideAnnotations.value.get(0)).toHaveLength(1);
	});

	it('clearAnnotations clears the active slide; clearAll clears everything', () => {
		const idx = ref(0);
		const a = usePresentationAnnotations({ isActive: ref(true), activeSlideIndex: idx });
		a.setPresentationTool('pen');
		a.handlePointerDown(0, 0);
		a.handlePointerMove(1, 1);
		a.handlePointerUp();

		a.clearAnnotations();
		expect(a.annotationStrokes.value).toHaveLength(0);
		expect(a.hasAnyAnnotations.value).toBeFalsy();

		// Draw again on two slides, then clear all.
		a.handlePointerDown(0, 0);
		a.handlePointerMove(1, 1);
		a.handlePointerUp();
		idx.value = 1;
		a.handlePointerDown(0, 0);
		a.handlePointerMove(1, 1);
		a.handlePointerUp();
		a.clearAllAnnotations();
		expect(a.hasAnyAnnotations.value).toBeFalsy();
		expect(a.allSlideAnnotations.value.size).toBe(0);
	});

	it('resets transient state when deactivated', () => {
		const active = ref(true);
		const a = usePresentationAnnotations({ isActive: active, activeSlideIndex: ref(0) });
		a.setPresentationTool('pen');
		a.handlePointerDown(0, 0);
		active.value = false;
		expect(a.presentationTool.value).toBe('none');
		expect(a.currentStroke.value).toBeNull();
	});
});

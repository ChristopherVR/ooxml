import { describe, expect, it } from 'vitest';

import { ViewerState } from './viewer-state.svelte';

describe('viewerState', () => {
	it('resets for a loaded deck with a clamped initial slide', () => {
		const state = new ViewerState();
		state.reset(5, 99);
		expect(state.slideCount).toBe(5);
		expect(state.current).toBe(4);
		state.reset(3);
		expect(state.current).toBe(0);
	});

	it('navigates with clamping', () => {
		const state = new ViewerState();
		state.reset(3);
		state.next();
		expect(state.current).toBe(1);
		state.last();
		expect(state.current).toBe(2);
		state.next();
		expect(state.current).toBe(2);
		state.first();
		state.prev();
		expect(state.current).toBe(0);
	});

	it('notes a slide gone to with its own selection for one slide change, clamped', () => {
		const state = new ViewerState();
		state.reset(3);
		state.goToWithSelection(9);
		expect(state.current).toBe(2);
		expect(state.takeSelectionSlide(1)).toBeFalsy();
		expect(state.takeSelectionSlide(2)).toBeFalsy();
		state.goToWithSelection(1);
		expect(state.takeSelectionSlide(1)).toBeTruthy();
		expect(state.takeSelectionSlide(1)).toBeFalsy();
		state.goTo(2);
		expect(state.takeSelectionSlide(2)).toBeFalsy();
	});

	it('keeps navigation bounds in sync with edited slide count', () => {
		const state = new ViewerState();
		state.reset(1);
		state.setSlideCount(2);
		state.goTo(1);
		expect(state.current).toBe(1);
		state.setSlideCount(1);
		expect(state.current).toBe(0);
	});

	it('handles navigation keys and reports handled-ness', () => {
		const state = new ViewerState();
		state.reset(4);
		expect(state.handleNavigationKey('ArrowRight')).toBeTruthy();
		expect(state.current).toBe(1);
		expect(state.handleNavigationKey('End')).toBeTruthy();
		expect(state.current).toBe(3);
		expect(state.handleNavigationKey('Home')).toBeTruthy();
		expect(state.current).toBe(0);
		expect(state.handleNavigationKey('x')).toBeFalsy();
	});

	it('ignores navigation keys before a deck is loaded', () => {
		const state = new ViewerState();
		expect(state.handleNavigationKey('ArrowRight')).toBeFalsy();
	});

	it('zooms from the effective percent and returns to fit', () => {
		const state = new ViewerState();
		expect(state.zoomPercent).toBeNull();
		state.zoomIn(100);
		expect(state.zoomPercent).toBe(110);
		state.zoomIn(999);
		expect(state.zoomPercent).toBe(120);
		state.zoomOut(999);
		expect(state.zoomPercent).toBe(110);
		state.zoomToFit();
		expect(state.zoomPercent).toBeNull();
	});

	it('never zooms out past fit, and stepping back down to 100% is fit', () => {
		const state = new ViewerState();
		state.zoomOut(100);
		expect(state.zoomPercent).toBeNull();
		state.zoomIn(100);
		state.zoomOut(110);
		expect(state.zoomPercent).toBeNull();
	});

	it('opens another slide at fit', () => {
		const state = new ViewerState();
		state.slideIdAt = (index) => `s${index}`;
		state.reset(5);
		// The viewer's effect reports the slide a load shows.
		state.showSlide('s0');
		state.zoomIn(100);
		state.zoomIn(110);
		expect(state.zoomPercent).toBe(120);
		state.goTo(3);
		expect(state.zoomPercent).toBeNull();

		state.zoomIn(100);
		state.goTo(3);
		expect(state.zoomPercent).toBe(110);

		// The host's goTo then setZoom, before the effect reports the slide.
		state.goTo(4);
		state.zoomPercent = 150;
		state.showSlide('s4');
		expect(state.zoomPercent).toBe(150);
	});
});

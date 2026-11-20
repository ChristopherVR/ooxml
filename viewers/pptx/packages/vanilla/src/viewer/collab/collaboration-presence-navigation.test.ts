/**
 * A broadcast's audience follows the presenter: each time the session owner moves to a slide, a
 * local viewer goes there too, and it may still look back at another slide in between.
 */
import type { CollaborationConfig, SanitizedPresence } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import type { CollaborationController } from './collaboration-controller-types';
import { syncPresenceNavigation } from './collaboration-presence-navigation';

const slides = Array.from({ length: 5 }, (_, i) => ({
	id: `s${i}`,
	elements: [],
})) as unknown as PptxSlide[];

function presenter(slide: number): SanitizedPresence[] {
	return [
		{
			clientId: 7,
			userName: 'Presenter',
			userColor: '#000000',
			role: 'owner',
			activeSlideIndex: slide,
			cursorX: 0,
			cursorY: 0,
		} as SanitizedPresence,
	];
}

function audience(role: CollaborationConfig['role']) {
	const store = createStore({ ...createInitialViewerState(), slides });
	const collaboration = {
		isActive: () => true,
		getConfig: () => ({ roomId: 'show', serverUrl: '', userName: 'A', role }),
		setActiveSlide: vi.fn(),
		setSelection: vi.fn(),
	} as unknown as CollaborationController;
	const goTo = vi.fn((index: number) => store.set({ currentSlide: index }));
	const stop = syncPresenceNavigation(store, collaboration, goTo);
	return { store, goTo, stop };
}

describe('broadcast audience navigation', () => {
	it('follows each move of the presenter', () => {
		const { store, stop } = audience('viewer');
		store.set({ remotePresences: presenter(2) });
		expect(store.get().currentSlide).toBe(2);
		store.set({ currentSlide: 0 });
		store.set({ remotePresences: presenter(2) });
		expect(store.get().currentSlide).toBe(0);
		store.set({ remotePresences: presenter(3) });
		expect(store.get().currentSlide).toBe(3);
		stop();
	});

	it('leaves a collaborator free to navigate', () => {
		const { store, goTo, stop } = audience('collaborator');
		store.set({ remotePresences: presenter(2) });
		expect(goTo).not.toHaveBeenCalled();
		stop();
	});
});

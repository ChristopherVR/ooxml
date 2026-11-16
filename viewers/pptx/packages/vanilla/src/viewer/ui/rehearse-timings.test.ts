import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createInitialViewerState, createStore } from '../state';
import { openRehearseTimings } from './rehearse-timings';

afterEach(() => {
	vi.useRealTimers();
	document.body.replaceChildren();
});

describe('openRehearseTimings (shared rehearsal controller)', () => {
	it('runs the show, records each slide and saves on request', async () => {
		vi.useFakeTimers();
		const host = document.createElement('div');
		document.body.appendChild(host);
		const store = createStore({ ...createInitialViewerState(), currentSlide: 0 });
		const onSave = vi.fn();
		const enterPresentation = vi.fn(async () => {
			store.set({ presenting: true });
		});
		const close = openRehearseTimings(document, host, createTranslator('en'), {
			store,
			enterPresentation,
			onSave,
		});
		await vi.advanceTimersByTimeAsync(0);
		expect(enterPresentation).toHaveBeenCalledOnce();
		await vi.advanceTimersByTimeAsync(2000);
		store.set({ currentSlide: 1 });
		await vi.advanceTimersByTimeAsync(1000);
		store.set({ presenting: false });
		const save = [...host.querySelectorAll('button')].find(
			(button) => button.textContent === 'Save Timings' || /save/i.test(button.textContent ?? ''),
		);
		expect(save).toBeTruthy();
		save!.click();
		expect(onSave).toHaveBeenCalledWith({ 0: 2000, 1: 1000 });
		expect(host.querySelector('.pptxv-rehearse')).toBeNull();
		close();
	});
});

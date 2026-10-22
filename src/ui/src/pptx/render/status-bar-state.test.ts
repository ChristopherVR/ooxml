import { describe, expect, it, vi } from 'vitest';

import { dispatchStatusBarIntent, statusBarState } from './status-bar-state';
import type { StatusBarControlId, StatusBarHostState } from './status-bar-state';

const translate = (key: string, params?: Record<string, string | number>) =>
	params ? `${key} ${JSON.stringify(params)}` : key;

const host = (over: Partial<StatusBarHostState> = {}): StatusBarHostState => ({
	slideCount: 3,
	activeSlideIndex: 1,
	dirty: false,
	zoomPercent: 120,
	translate,
	...over,
});

describe('statusBarState', () => {
	it('shows every control a host handles, with the saved state', () => {
		expect(statusBarState(host({ notesExpanded: true, viewMode: 'normal' }))).toStrictEqual({
			slideCount: 3,
			activeSlideIndex: 1,
			saveText: 'pptx.statusBar.allSaved',
			saveKind: 'idle',
			zoomPercent: 120,
			showNotes: true,
			notesExpanded: true,
			showViewModes: true,
			showSorter: true,
			showSlideShow: true,
			viewMode: 'normal',
			translate,
		});
	});

	it('reads the autosave time, then the dirty flag', () => {
		const now = Date.UTC(2026, 0, 1, 12, 5);
		const saved = { state: 'saved' as const, timestamp: now - 3 * 60_000 };
		expect(statusBarState(host({ autosave: saved, now })).saveText).toBe(
			'pptx.autosave.saved {"time":"pptx.autosave.minutesAgo {\\"count\\":3}"}',
		);
		// Without its time a "saved" state falls back to the dirty flag.
		expect(statusBarState(host({ autosave: { state: 'saved' }, dirty: true })).saveText).toBe(
			'pptx.statusBar.unsavedChanges',
		);
		expect(statusBarState(host({ autosave: { state: 'error' } }))).toMatchObject({
			saveText: 'pptx.autosave.error',
			saveKind: 'error',
		});
	});

	it('hides a control the host cannot handle or has hidden', () => {
		const state = statusBarState(
			host({
				handles: { notes: false, sorter: false },
				hiddenActions: ['zoom', 'fullscreen'],
			}),
		);
		expect(state).toMatchObject({
			showNotes: false,
			showViewModes: true,
			showSorter: false,
			showSlideShow: false,
		});
		expect(state.zoomPercent).toBeUndefined();
		expect(statusBarState(host({ hiddenActions: ['notes'] })).showNotes).toBeFalsy();
		expect(statusBarState(host({ handles: { viewModes: false } })).showViewModes).toBeFalsy();
		expect(statusBarState(host({ zoomPercent: undefined })).zoomPercent).toBeUndefined();
		expect(statusBarState(host())).not.toHaveProperty('viewMode');
	});
});

describe('dispatchStatusBarIntent', () => {
	const request = (id: StatusBarControlId) => new CustomEvent('status-request', { detail: { id } });

	it('calls the handler for the requested control', () => {
		const handlers = { notes: vi.fn(), zoomIn: vi.fn() };
		dispatchStatusBarIntent(request('zoomIn'), handlers);
		expect(handlers.zoomIn).toHaveBeenCalledOnce();
		expect(handlers.notes).not.toHaveBeenCalled();
	});

	it('ignores a control without a handler and an event without an intent', () => {
		const handlers = { notes: vi.fn() };
		expect(() => dispatchStatusBarIntent(request('sorter'), handlers)).not.toThrow();
		expect(() => dispatchStatusBarIntent(new Event('status-request'), handlers)).not.toThrow();
		expect(handlers.notes).not.toHaveBeenCalled();
	});
});

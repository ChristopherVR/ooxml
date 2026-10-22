// @vitest-environment happy-dom
import type { ViewerOptions, ViewerOptionsStore } from 'ooxml-ui/pptx';
import {
	DEFAULT_VIEWER_OPTIONS,
	VIEWER_PREFS_STORAGE_KEY,
	clearAutosaveSnapshots,
	mergeViewerOptions,
	pruneExpiredAutosaveSnapshots,
} from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { useAutosaveCacheMaintenance } from './useAutosaveCacheMaintenance';
import { useViewerOptions } from './useViewerOptions';

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal<typeof import('ooxml-ui/pptx')>()),
	// The sweeps themselves (expiry, host recovery stores) are covered by the shared
	// `autosave-maintenance` tests; here only when the hook runs them matters.
	pruneExpiredAutosaveSnapshots: vi.fn(),
	clearAutosaveSnapshots: vi.fn(),
}));

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	vi.mocked(pruneExpiredAutosaveSnapshots).mockReset().mockResolvedValue(0);
	vi.mocked(clearAutosaveSnapshots).mockReset().mockResolvedValue(0);
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	localStorage.clear();
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(async () => {
	await act(async () => root.unmount());
	container.remove();
	localStorage.clear();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
	vi.restoreAllMocks();
});

interface Captured {
	store: ViewerOptionsStore;
	options: ViewerOptions;
}

function mountHook(): { current: () => Captured } {
	const captured: { value: Captured | null } = { value: null };
	function Probe(): null {
		const { optionsStore, options } = useViewerOptions();
		captured.value = { store: optionsStore, options };
		return null;
	}
	act(() => {
		root.render(React.createElement(Probe));
	});
	return {
		current: () => {
			expect(captured.value).not.toBeNull();
			return captured.value as Captured;
		},
	};
}

function readStoredOptions(): Record<string, Record<string, unknown>> {
	const raw = localStorage.getItem(VIEWER_PREFS_STORAGE_KEY);
	expect(raw).not.toBeNull();
	const parsed = JSON.parse(raw as string) as { options?: Record<string, Record<string, unknown>> };
	return parsed.options ?? {};
}

describe('useViewerOptions', () => {
	it('starts from the shared defaults', () => {
		const hook = mountHook();
		expect(hook.current().options).toStrictEqual(DEFAULT_VIEWER_OPTIONS);
	});

	it('reflects store changes in the reactive snapshot', () => {
		const hook = mountHook();
		act(() => {
			hook.current().store.setValue('advanced', 'maximumUndoSteps', 42);
		});
		expect(hook.current().options.advanced.maximumUndoSteps).toBe(42);
		// Untouched groups keep their default values.
		expect(hook.current().options.general.showMiniToolbar).toBeTruthy();
	});

	it('persists a sparse diff to localStorage', () => {
		const hook = mountHook();
		act(() => {
			hook.current().store.setValue('proofing', 'autoCorrectSmartQuotes', false);
			hook.current().store.setRibbonTabHidden('review', true);
		});
		const stored = readStoredOptions();
		expect(stored.proofing?.autoCorrectSmartQuotes).toBeFalsy();
		expect(stored.ribbon?.hiddenTabIds).toStrictEqual(['review']);
		// Defaults are not persisted (sparse diff).
		expect(stored.general).toBeUndefined();
	});

	it('hydrates persisted values in a fresh hook instance', () => {
		const first = mountHook();
		act(() => {
			first.current().store.setValue('save', 'autoRecoverIntervalMinutes', 7);
		});
		act(() => root.unmount());

		root = createRoot(container);
		const second = mountHook();
		expect(second.current().options.save.autoRecoverIntervalMinutes).toBe(7);
	});

	it('does not add cache maintenance to a headless options store', async () => {
		const add = vi.spyOn(window, 'addEventListener');
		const hook = mountHook();
		await act(async () => {
			hook.current().store.setValue('save', 'clearCacheOnClose', true);
			window.dispatchEvent(new Event('beforeunload'));
			root.render(null);
		});
		expect(pruneExpiredAutosaveSnapshots).not.toHaveBeenCalled();
		expect(clearAutosaveSnapshots).not.toHaveBeenCalled();
		expect(add.mock.calls.filter(([name]) => name === 'beforeunload')).toHaveLength(0);
	});
});

function mountMaintenance(initial = DEFAULT_VIEWER_OPTIONS) {
	let clear: () => void;
	function Probe({ options }: { options: ViewerOptions }): null {
		clear = useAutosaveCacheMaintenance(options);
		return null;
	}
	const render = (options: ViewerOptions): void => {
		act(() => root.render(React.createElement(Probe, { options })));
	};
	render(initial);
	return { render, clear: () => clear };
}

describe('full-viewer autosave cache maintenance', () => {
	it('sweeps once using the mount options even if options change', async () => {
		const mounted = mergeViewerOptions({ save: { cacheRetentionDays: 7 } });
		const hook = mountMaintenance(mounted);
		hook.render(mergeViewerOptions({ save: { cacheRetentionDays: 30 } }));
		await act(async () => {});
		expect(pruneExpiredAutosaveSnapshots).toHaveBeenCalledExactlyOnceWith(mounted, undefined);
	});

	it('returns the same module-level clear callback across renders and mounts', () => {
		const first = mountMaintenance();
		const clear = first.clear();
		first.render(mergeViewerOptions({ save: { cacheRetentionDays: 30 } }));
		expect(first.clear()).toBe(clear);
		act(() => root.render(null));
		expect(mountMaintenance().clear()).toBe(clear);
	});

	it('manually clears every snapshot regardless of retention or close preferences', async () => {
		const hook = mountMaintenance();
		await act(async () => {
			expect(hook.clear()()).toBeUndefined();
		});
		expect(clearAutosaveSnapshots).toHaveBeenCalledExactlyOnceWith(undefined);
	});

	it('hands a host storage to both sweeps', async () => {
		const storage = { save: async () => undefined };
		let clear!: () => void;
		function Probe(): null {
			clear = useAutosaveCacheMaintenance(DEFAULT_VIEWER_OPTIONS, storage);
			return null;
		}
		await act(async () => root.render(React.createElement(Probe)));
		await act(async () => clear());
		expect(pruneExpiredAutosaveSnapshots).toHaveBeenCalledWith(DEFAULT_VIEWER_OPTIONS, storage);
		expect(clearAutosaveSnapshots).toHaveBeenCalledWith(storage);
	});

	it('removes the exact listener on option changes and unmount without clearing by default', async () => {
		const add = vi.spyOn(window, 'addEventListener');
		const remove = vi.spyOn(window, 'removeEventListener');
		const hook = mountMaintenance();
		hook.render(mergeViewerOptions({ general: { showMiniToolbar: false } }));
		await act(async () => {
			window.dispatchEvent(new Event('beforeunload'));
			root.render(null);
		});
		const added = add.mock.calls.filter(([name]) => name === 'beforeunload');
		const removed = remove.mock.calls.filter(([name]) => name === 'beforeunload');
		expect(added).toHaveLength(2);
		expect(removed).toStrictEqual(added);
		expect(pruneExpiredAutosaveSnapshots).toHaveBeenCalledOnce();
		expect(clearAutosaveSnapshots).not.toHaveBeenCalled();
	});

	it('preserves clearing on requested close and option-effect cleanup', async () => {
		const enabled = mergeViewerOptions({ save: { clearCacheOnClose: true } });
		const hook = mountMaintenance(enabled);
		await act(async () => {
			window.dispatchEvent(new Event('beforeunload'));
		});
		expect(clearAutosaveSnapshots).toHaveBeenCalledOnce();
		await act(async () => hook.render(DEFAULT_VIEWER_OPTIONS));
		expect(clearAutosaveSnapshots).toHaveBeenCalledTimes(2);
		await act(async () => {
			window.dispatchEvent(new Event('beforeunload'));
		});
		expect(clearAutosaveSnapshots).toHaveBeenCalledTimes(2);
		await act(async () => hook.render(enabled));
		await act(async () => root.render(null));
		expect(clearAutosaveSnapshots).toHaveBeenCalledTimes(3);
		await act(async () => {
			window.dispatchEvent(new Event('beforeunload'));
		});
		expect(clearAutosaveSnapshots).toHaveBeenCalledTimes(3);
	});

	it.each(['prune', 'clear'] as const)('ignores background %s failures', async (operation) => {
		vi.mocked(
			operation === 'prune' ? pruneExpiredAutosaveSnapshots : clearAutosaveSnapshots,
		).mockRejectedValueOnce(new Error('Blocked storage'));
		const hook = await act(async () => mountMaintenance());
		await act(async () => hook.clear()());
		expect(pruneExpiredAutosaveSnapshots).toHaveBeenCalledOnce();
	});
});

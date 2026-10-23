import type { ConnectionStatus, ViewerAutosaveStatus } from 'ooxml-ui/pptx';
import type { PptxCoreProperties } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ViewerStateBag } from './create-viewer-state-types';
import CreateViewerStateHarness from './CreateViewerStateHarness.svelte';

/**
 * The host events every binding reports the same way (issue 151, item 5), driven through the
 * real `createViewerState` factory mounted in a component.
 */

let cleanup: (() => void) | undefined;
beforeEach(() => {
	localStorage.clear();
});
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

interface HarnessProps {
	editable?: boolean;
	filePath?: string;
	onpropertieschange?: (updates: Partial<PptxCoreProperties>) => void;
	onpresentationchange?: (presenting: boolean) => void;
	onautosavestatus?: (status: ViewerAutosaveStatus) => void;
	oncollaborationstatus?: (status: ConnectionStatus) => void;
}

function renderHarness(props: HarnessProps = {}): ViewerStateBag {
	let captured: ViewerStateBag | undefined;
	const instance = mount(CreateViewerStateHarness, {
		target: document.createElement('div'),
		props: {
			...props,
			onready: (state: ViewerStateBag) => {
				captured = state;
			},
		},
	});
	cleanup = () => unmount(instance);
	if (!captured) {
		throw new Error('createViewerState harness did not report its state synchronously');
	}
	flushSync();
	return captured;
}

describe('onpresentationchange', () => {
	it('reports the slide show starting and stopping, not the initial state', () => {
		const onpresentationchange = vi.fn();
		const state = renderHarness({ onpresentationchange });
		expect(onpresentationchange).not.toHaveBeenCalled();

		state.viewer.isFullscreen = true;
		flushSync();
		expect(onpresentationchange).toHaveBeenLastCalledWith(true);

		state.viewer.isFullscreen = false;
		flushSync();
		expect(onpresentationchange).toHaveBeenLastCalledWith(false);
		expect(onpresentationchange).toHaveBeenCalledTimes(2);
	});
});

describe('onautosavestatus', () => {
	it('reports the public lifecycle of the shared engine, not the initial state', async () => {
		const onautosavestatus = vi.fn();
		const state = renderHarness({ onautosavestatus, filePath: 'deck.pptx' });
		expect(onautosavestatus).not.toHaveBeenCalled();

		// No deck is loaded, so the write is skipped: saving, then back to idle.
		const pending = state.autosaveCtl.save();
		flushSync();
		await pending;
		flushSync();

		expect(onautosavestatus.mock.calls).toStrictEqual([['saving'], ['idle']]);
	});
});

describe('oncollaborationstatus', () => {
	it('reports each connection-status transition', () => {
		const oncollaborationstatus = vi.fn();
		const state = renderHarness({ oncollaborationstatus });
		expect(oncollaborationstatus).not.toHaveBeenCalled();

		state.collab.status = 'connecting';
		flushSync();
		state.collab.status = 'connected';
		flushSync();

		expect(oncollaborationstatus.mock.calls).toStrictEqual([['connecting'], ['connected']]);
	});
});

describe('onpropertieschange', () => {
	it('reports a user edit of the core properties', () => {
		const onpropertieschange = vi.fn();
		const state = renderHarness({ editable: true, onpropertieschange });

		state.editor.updateDocumentProperties({ title: 'Quarterly review' }, {}, []);

		expect(onpropertieschange).toHaveBeenCalledExactlyOnceWith({ title: 'Quarterly review' });
		expect(state.editor.coreProperties?.title).toBe('Quarterly review');
	});

	it('does not report an edit that is not a user change to the core properties', () => {
		const onpropertieschange = vi.fn();
		const state = renderHarness({ editable: true, onpropertieschange });

		state.editor.updateDocumentProperties({ title: 'From the assistant' }, {}, [], null);

		expect(onpropertieschange).not.toHaveBeenCalled();
	});

	it('stays silent for a read-only viewer', () => {
		const onpropertieschange = vi.fn();
		const state = renderHarness({ editable: false, onpropertieschange });

		state.editor.updateDocumentProperties({ title: 'Nope' }, {}, []);

		expect(onpropertieschange).not.toHaveBeenCalled();
	});
});

import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createInitialViewerState, createStore } from '../state';
import type { ViewerChrome } from '../ui';
import { createCollabUi } from './collab-ui';
import type { CollabUiDeps } from './collab-ui';

/** A minimal fake chrome exposing only what `collab-ui.ts` touches. */
function fakeChrome(): ViewerChrome {
	const ribbonEl = document.createElement('div');
	const primaryRow = document.createElement('div');
	primaryRow.className = 'pptxv-ribbon-primary';
	ribbonEl.appendChild(primaryRow);
	const collaborationHost = document.createElement('span');
	return {
		root: document.createElement('div'),
		ribbon: { el: ribbonEl } as unknown as ViewerChrome['ribbon'],
		mobileToolbar: { collaborationHost } as unknown as ViewerChrome['mobileToolbar'],
		stageWrap: document.createElement('div'),
	} as unknown as ViewerChrome;
}

function makeDeps(over: Partial<CollabUiDeps> = {}): CollabUiDeps {
	const store = createStore(createInitialViewerState());
	const chrome = fakeChrome();
	return {
		doc: document,
		store,
		getChrome: () => chrome,
		getTranslator: () => createTranslator(),
		getScale: () => 1,
		startCollaboration: vi.fn(() => Promise.resolve()),
		stopCollaboration: vi.fn(),
		getStatus: () => 'disconnected',
		getConfig: () => null,
		followUser: vi.fn(),
		...over,
	};
}

describe('createCollabUi', () => {
	it('omitting hiddenActions renders only the Share quick-access button (Broadcast lives in the Present menu, matching React)', () => {
		const deps = makeDeps();
		createCollabUi(deps);
		const primaryRow = deps.getChrome().ribbon?.el.querySelector('.pptxv-ribbon-primary');
		expect(primaryRow?.querySelectorAll(':scope > button')).toHaveLength(1);
	});

	it("hides the Share quick-access button (desktop + mobile) on 'share'", () => {
		const deps = makeDeps({ hiddenActions: ['share'] });
		createCollabUi(deps);
		const primaryRow = deps.getChrome().ribbon?.el.querySelector('.pptxv-ribbon-primary');
		expect(primaryRow?.querySelectorAll(':scope > button')).toHaveLength(0);
		const mobileHost = deps.getChrome().mobileToolbar?.collaborationHost;
		expect(mobileHost?.querySelector('button')).toBeNull();
	});

	it('mounts the overlays in the stage wrap and the follow bar on the viewer root', () => {
		const deps = makeDeps();
		createCollabUi(deps);
		const chrome = deps.getChrome();
		expect(chrome.stageWrap.querySelector('.pptxv-collab-cursors')).not.toBeNull();
		expect(chrome.stageWrap.querySelector('.pptxv-remote-selections')).not.toBeNull();
		// The follow bar is a viewer-viewport pill, not a stage overlay.
		expect(chrome.stageWrap.querySelector('.pptxv-follow-bar')).toBeNull();
		expect(chrome.root.querySelector('.pptxv-follow-bar')).not.toBeNull();
	});

	it("'broadcast' in hiddenActions does not remove the Share quick-access button", () => {
		const deps = makeDeps({ hiddenActions: ['broadcast'] });
		createCollabUi(deps);
		const primaryRow = deps.getChrome().ribbon?.el.querySelector('.pptxv-ribbon-primary');
		expect(primaryRow?.querySelectorAll(':scope > button')).toHaveLength(1);
		const mobileHost = deps.getChrome().mobileToolbar?.collaborationHost;
		expect(mobileHost?.querySelector('button')).not.toBeNull();
	});
});

describe('createCollabUi host events', () => {
	afterEach(() => {
		document.body.replaceChildren();
	});

	/** Open the broadcast dialog and press its Start action (a `dialog-footer-request`). */
	function startBroadcastFromDialog(deps: CollabUiDeps): void {
		createCollabUi(deps).openBroadcast();
		// The Share dialog mounts first, so the Broadcast dialog's footer is the last one.
		const footer = [...document.body.querySelectorAll('pptx-ui-dialog-footer')].at(-1);
		expect(footer).toBeDefined();
		// The dialog opens with the (blank) remembered room; give it one to start.
		let scope: HTMLElement | null = footer?.parentElement ?? null;
		while (scope && !scope.querySelector('input')) {
			scope = scope.parentElement;
		}
		const room = scope?.querySelector('input');
		expect(room).toBeDefined();
		if (room) {
			room.value = 'room-1';
			room.dispatchEvent(new Event('input', { bubbles: true }));
		}
		footer?.dispatchEvent(
			new CustomEvent('dialog-footer-request', { detail: { id: 'start' }, bubbles: true }),
		);
	}

	it('reports the started session, named after authorName, to onStarted', () => {
		const onStarted = vi.fn();
		const deps = makeDeps({ getAuthorName: () => 'Dana Host', onStarted });

		startBroadcastFromDialog(deps);

		expect(onStarted).toHaveBeenCalledWith(
			expect.objectContaining({ userName: 'Dana Host', role: 'owner' }),
		);
		expect(deps.startCollaboration).toHaveBeenCalledTimes(1);
	});

	it('reports a stop from the dialog to onStopped', () => {
		const onStopped = vi.fn();
		const deps = makeDeps({ onStopped });
		startBroadcastFromDialog(deps);

		document.body.querySelector<HTMLButtonElement>('.pptxv-modal-danger-btn')?.click();

		expect(onStopped).toHaveBeenCalledTimes(1);
		expect(deps.stopCollaboration).toHaveBeenCalledTimes(1);
	});
});

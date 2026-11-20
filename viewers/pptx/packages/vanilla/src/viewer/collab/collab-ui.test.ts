import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createInitialViewerState, createStore } from '../state';
import type { ViewerChrome } from '../ui';
import { createCollabUi } from './collab-ui';
import type { CollabUiDeps } from './collab-ui';
import { createCollaborationController } from './collaboration-controller';

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
		getFailure: () => null,
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
		// Name the room, so the started session is a known one.
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

	it('opens a new broadcast with a fresh room, so Start works at once', () => {
		const deps = makeDeps();
		createCollabUi(deps).openBroadcast();
		const footer = [...document.body.querySelectorAll('pptx-ui-dialog-footer')].at(-1);
		const rooms = [...document.body.querySelectorAll<HTMLInputElement>('input')]
			.map((input) => input.value)
			.filter((value) => value.startsWith('broadcast-'));
		expect(rooms).toHaveLength(1);

		footer?.dispatchEvent(
			new CustomEvent('dialog-footer-request', { detail: { id: 'start' }, bubbles: true }),
		);
		expect(deps.startCollaboration).toHaveBeenCalledWith(
			expect.objectContaining({ roomId: rooms[0], role: 'owner', transport: 'webrtc' }),
		);
	});
});

describe('createCollabUi start failure', () => {
	afterEach(() => {
		document.body.replaceChildren();
	});

	it('shows why a peer-to-peer session without its secret could not start', async () => {
		const missingSecret = translationsEn['pptx.collaboration.failure.missingSecret'];
		const deps = makeDeps();
		const collab = createCollaborationController({
			store: deps.store,
			getHandler: () => null,
			setEditable: () => {},
		});
		const ui = createCollabUi({
			...deps,
			startCollaboration: (config) => collab.start(config),
			stopCollaboration: () => collab.stop(),
			getStatus: () => collab.getStatus(),
			getFailure: () => collab.getFailure(),
			getConfig: () => collab.getConfig(),
		});
		await collab.start({ roomId: 'room-1', serverUrl: '', userName: 'Ada', transport: 'webrtc' });
		ui.onStatusChange(collab.getStatus());
		expect(collab.getStatus()).toBe('error');

		const pill = deps.getChrome().ribbon?.el.querySelector('[data-testid="collaboration-status"]');
		expect(pill?.textContent).toContain(translationsEn['pptx.collaboration.status.error']);
		expect(pill?.getAttribute('title')).toBe(missingSecret);

		ui.openShare();
		const banner = document.body.querySelector<HTMLElement>(
			'[data-testid="collaboration-failure"]',
		);
		expect(banner?.hidden).toBeFalsy();
		expect(banner?.textContent).toBe(missingSecret);
		collab.destroy();
		ui.destroy();
	});

	it('says why a peer-to-peer invitation without its key cannot join', () => {
		const missingSecret = translationsEn['pptx.collaboration.failure.missingSecret'];
		const ui = createCollabUi(makeDeps());
		ui.openShare();
		document.body.querySelectorAll<HTMLButtonElement>('.pptxv-share-tabs button')[1]?.click();
		const field = Array.from(document.body.querySelectorAll<HTMLInputElement>('input')).find(
			(input) => input.placeholder === translationsEn['pptx.share.invitationPlaceholder'],
		);
		field!.value = 'https://x.test/?room=deck-1&transport=webrtc';
		field!.dispatchEvent(new Event('input', { bubbles: true }));
		const banner = document.body.querySelector<HTMLElement>(
			'[data-testid="collaboration-failure"]',
		);
		expect(banner?.hidden).toBeFalsy();
		expect(banner?.textContent).toBe(missingSecret);
		ui.destroy();
	});
});

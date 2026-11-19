/**
 * A peer-to-peer session that cannot start says why: the real
 * `CollaborationController` runs the shared session controller against a
 * webrtc config that lacks its invitation secret, and the status pill and the
 * Share dialog show the translated reason instead of only "Connection error".
 */
import type { CollaborationConfig } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { CollaborationController } from '../collaboration.svelte';
import CollaborationStatusIndicator from './CollaborationStatusIndicator.svelte';
import ShareDialog from './ShareDialog.svelte';

const MISSING_SECRET = translationsEn['pptx.collaboration.failure.missingSecret'];

const cleanup: (() => void)[] = [];
afterEach(() => {
	for (const dispose of cleanup.splice(0).reverse()) {
		dispose();
	}
});

function mounted(component: unknown, props: Record<string, unknown>): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(component as never, { target, props } as never);
	flushSync();
	cleanup.push(() => {
		unmount(instance);
		target.remove();
	});
	return target;
}

describe('collaboration start failure (Svelte)', () => {
	it('shows why a peer-to-peer session without its secret could not start', async () => {
		const config: CollaborationConfig = {
			roomId: 'room-1',
			serverUrl: '',
			userName: 'Ada',
			transport: 'webrtc',
		};
		let collab: CollaborationController | undefined;
		cleanup.push(
			$effect.root(() => {
				collab = new CollaborationController({
					getSlides: () => [],
					applyRemoteSlides: () => {},
					getConfig: () => config,
				});
			}),
		);
		await collab!.start(config);
		expect(collab!.status).toBe('error');
		expect(collab!.failure).toBe('missing-secret');

		mounted(ShareDialog, {
			open: true,
			active: collab!.active,
			status: collab!.status,
			failure: collab!.failure,
			connectedCount: 0,
			remotePresences: [],
			activeCollaboration: null,
			onstart: () => {},
			onstop: () => {},
			onclose: () => {},
		});
		const banner = document.querySelector('[data-testid="collaboration-failure"]');
		expect(banner?.textContent).toBe(MISSING_SECRET);

		const pill = mounted(CollaborationStatusIndicator, {
			status: collab!.status,
			connectedCount: 0,
			failure: collab!.failure,
			onretry: () => {},
		});
		const status = pill.querySelector('[role="status"]');
		expect(status?.textContent).toContain(translationsEn['pptx.collaboration.status.error']);
		expect(status?.getAttribute('title')).toBe(MISSING_SECRET);
	});

	it('says why a peer-to-peer invitation without its key cannot join', () => {
		const dialog = mounted(ShareDialog, {
			open: true,
			defaults: { userName: 'Ada' },
			active: false,
			status: 'disconnected',
			failure: null,
			connectedCount: 0,
			remotePresences: [],
			activeCollaboration: null,
			onstart: () => {},
			onstop: () => {},
			onclose: () => {},
		});
		document.querySelectorAll<HTMLButtonElement>('[role="tab"]')[1]?.click();
		flushSync();
		const field = document.querySelector<HTMLInputElement>('#pptx-svelte-share-invitation');
		field!.value = 'https://x.test/?room=deck-1&transport=webrtc';
		field!.dispatchEvent(new Event('input', { bubbles: true }));
		flushSync();
		expect(
			dialog.ownerDocument.querySelector('[data-testid="collaboration-failure"]')?.textContent,
		).toBe(MISSING_SECRET);
	});
});

/**
 * A peer-to-peer session that cannot start says why: the real
 * `useCollaboration` runs the shared session controller against a webrtc
 * config that lacks its invitation secret, and the status pill and the Share
 * dialog show the translated reason instead of only "Connection error".
 */
import { mount } from '@vue/test-utils';
import type { CollaborationConfig } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { afterEach, describe, expect, it } from 'vitest';
import { effectScope, ref } from 'vue';

import { useCollaboration } from '../composables/useCollaboration';
import CollaborationStatusIndicator from './CollaborationStatusIndicator.vue';
import ShareDialog from './ShareDialog.vue';

const MISSING_SECRET = translationsEn['pptx.collaboration.failure.missingSecret'];

const cleanup: (() => void)[] = [];
afterEach(() => {
	for (const dispose of cleanup.splice(0).reverse()) {
		dispose();
	}
});

async function waitFor(check: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 100 && !check(); attempt += 1) {
		await new Promise<void>((resolve) => {
			setTimeout(resolve, 10);
		});
	}
}

describe('collaboration start failure (Vue)', () => {
	it('shows why a peer-to-peer session without its secret could not start', async () => {
		const config: CollaborationConfig = {
			roomId: 'room-1',
			serverUrl: '',
			userName: 'Ada',
			transport: 'webrtc',
		};
		const scope = effectScope();
		cleanup.push(() => scope.stop());
		const collab = scope.run(() =>
			useCollaboration({
				collaboration: config,
				slides: ref([]),
				onRemoteSlides: () => {},
				loadVersion: ref(0),
				getLoadOrigin: () => 'bootstrap',
			}),
		)!;
		await waitFor(() => collab.status.value === 'error');
		expect(collab.failure.value).toBe('missing-secret');

		const dialog = mount(ShareDialog, {
			global: { stubs: { teleport: true } },
			props: { open: true, active: collab.active.value, collab },
		});
		cleanup.push(() => dialog.unmount());
		expect(dialog.get('[data-testid="collaboration-failure"]').text()).toBe(MISSING_SECRET);

		const pill = mount(CollaborationStatusIndicator, {
			props: {
				status: collab.status.value,
				connectedCount: collab.connectedCount.value,
				failure: collab.failure.value,
			},
		});
		cleanup.push(() => pill.unmount());
		expect(pill.text()).toContain(translationsEn['pptx.collaboration.status.error']);
		expect(pill.attributes('title')).toBe(MISSING_SECRET);
	});

	it('says why a peer-to-peer invitation without its key cannot join', async () => {
		const dialog = mount(ShareDialog, {
			global: { stubs: { teleport: true } },
			props: { open: true, defaults: { userName: 'Ada' } },
		});
		cleanup.push(() => dialog.unmount());
		await dialog.findAll('[role="tab"]')[1]!.trigger('click');
		await dialog
			.get('#pptx-vue-share-invitation')
			.setValue('https://x.test/?room=deck-1&transport=webrtc');
		expect(dialog.get('[data-testid="collaboration-failure"]').text()).toBe(MISSING_SECRET);
	});
});

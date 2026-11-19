import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { flushPromises, mount } from '@vue/test-utils';
import type { CollaborationConfig } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';

const DECK = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

/**
 * A peer-to-peer link without its key never opens a session, so the status bar shows why on the
 * page, without the user opening Share first.
 */
describe('powerPointViewer collaboration status', () => {
	it('shows why a keyless peer-to-peer link cannot join in the status bar', async () => {
		const collaboration: CollaborationConfig = {
			roomId: 'room-1',
			serverUrl: '',
			userName: 'Ada',
			transport: 'webrtc',
		};
		const wrapper = mount(PowerPointViewer, {
			props: { content: new Uint8Array(readFileSync(DECK)).buffer, collaboration },
		});
		try {
			await vi.waitFor(
				async () => {
					await flushPromises();
					const pill = wrapper.find('[data-testid="collaboration-status"]');
					expect(pill.exists()).toBe(true);
					expect(pill.attributes('title')).toBe(
						translationsEn['pptx.collaboration.failure.missingSecret'],
					);
					expect(pill.text()).toContain(translationsEn['pptx.collaboration.status.error']);
				},
				{ timeout: 15_000 },
			);
		} finally {
			wrapper.unmount();
		}
	}, 30_000);
});

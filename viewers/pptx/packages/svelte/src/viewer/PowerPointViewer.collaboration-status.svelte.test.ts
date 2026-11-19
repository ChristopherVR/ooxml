import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { CollaborationConfig } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.svelte';

const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

/**
 * A peer-to-peer link without its key never opens a session, so the status bar shows why on the
 * page, without the user opening Share first.
 */
describe('svelte viewer collaboration status', () => {
	it('shows why a keyless peer-to-peer link cannot join in the status bar', async () => {
		const collaboration: CollaborationConfig = {
			roomId: 'room-1',
			serverUrl: '',
			userName: 'Ada',
			transport: 'webrtc',
		};
		const target = document.createElement('div');
		document.body.appendChild(target);
		const instance = mount(PowerPointViewer, {
			target,
			props: { source: new Uint8Array(readFileSync(FIXTURE)), collaboration },
		});
		cleanup = () => {
			unmount(instance);
			target.remove();
		};
		await vi.waitFor(
			() => {
				flushSync();
				const pill = target.querySelector('[data-testid="collaboration-status"]');
				expect(pill?.getAttribute('title')).toBe(
					translationsEn['pptx.collaboration.failure.missingSecret'],
				);
				expect(pill?.textContent).toContain(translationsEn['pptx.collaboration.status.error']);
			},
			{ timeout: 15_000 },
		);
	}, 30_000);
});

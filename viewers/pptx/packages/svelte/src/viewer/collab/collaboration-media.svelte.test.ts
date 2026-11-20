/**
 * The room carries the deck's media, and linked content follows this viewer's own setting:
 * the controller hands the room the URLs the viewer plays its video from, and reads the
 * room's slides with the viewer's remote content setting. `.svelte.test.ts` so the runes
 * runtime compiles the controller's `$effect`s.
 */
import type { AwarenessLike, CollaborationConfig, YDocLike, YjsFactories } from 'ooxml-ui/pptx';
import { reconcileSlidesInYDoc } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import type { CollabSession, CollabSessionFactory } from './collaboration-session';
import { CollaborationController } from './collaboration.svelte';

const CONFIG: CollaborationConfig = {
	roomId: 'media-room',
	serverUrl: '',
	transport: 'webrtc',
	userName: 'Tester',
};
const VIDEO_PATH = 'ppt/media/media1.mp4';
const PLAYED_FROM = 'blob:http://localhost/played-video';
const LINKED = 'https://linked.test/picture.png';

const factories = (): YjsFactories => ({
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
});
const deck = (...elements: unknown[]): PptxSlide[] =>
	[{ id: 's1', rId: 'rId2', slideNumber: 1, elements }] as unknown as PptxSlide[];

function awareness(): AwarenessLike {
	const states = new Map<number, Record<string, unknown>>();
	return {
		clientID: 1,
		setLocalStateField: (field, value) => states.set(1, { [field]: value }),
		getStates: () => states,
		on: () => {},
		off: () => {},
	};
}

function sessionOn(doc: Y.Doc): CollabSessionFactory {
	return async (): Promise<CollabSession> => ({
		ydoc: doc as unknown as YDocLike,
		factories: factories(),
		provider: {
			awareness: awareness(),
			onStatus: () => {},
			connectedNow: true,
			onSynced: () => {},
			syncedNow: true,
			destroy: vi.fn(),
		},
		destroy: vi.fn(),
	});
}

/** Run `body` inside a live `$effect.root` until its assertions settle. */
function inRoot(body: () => Promise<void>): Promise<void> {
	return new Promise<void>((resolve, reject) => {
		let dispose = (): void => {};
		dispose = $effect.root(() => {
			void (async () => {
				try {
					await body();
					resolve();
				} catch (err) {
					reject(err instanceof Error ? err : new Error(String(err)));
				} finally {
					dispose();
				}
			})();
		});
	});
}

describe('collaboration media in the Svelte binding', () => {
	it('shares an embedded video as a media token, never the URL it plays from', async () => {
		const doc = new Y.Doc();
		await inRoot(async () => {
			const slides = deck({ id: 'vid', type: 'media', mediaType: 'video', mediaPath: VIDEO_PATH });
			const collab = new CollaborationController({
				getSlides: () => slides,
				applyRemoteSlides: () => {},
				getConfig: () => CONFIG,
				getMediaDataUrls: () => new Map([[VIDEO_PATH, PLAYED_FROM]]),
				createSession: sessionOn(doc),
			});
			await collab.start(CONFIG);
			await vi.waitFor(() => expect(doc.getArray('pptx:slides')).toHaveLength(1));
			const assets = doc.getMap('pptx:assets').toJSON() as Record<string, string>;
			expect(assets['vid:mediaData']).toMatch(/^ooxml-media:[0-9a-f]{32}$/u);
			expect(JSON.stringify(doc.toJSON())).not.toContain('blob:');
			collab.stop();
		});
		doc.destroy();
	});

	it.each([
		[false, undefined],
		[true, LINKED],
	])(
		'with external content allowed: %s, shows a linked picture from the room as %s',
		async (allow, expected) => {
			const doc = new Y.Doc();
			reconcileSlidesInYDoc(
				deck({ id: 'pic', type: 'picture', imagePath: LINKED, imageData: LINKED }),
				doc as unknown as YDocLike,
				factories(),
				'remote',
			);
			await inRoot(async () => {
				let received: PptxSlide[] = [];
				const collab = new CollaborationController({
					getSlides: () => deck(),
					applyRemoteSlides: (slides) => {
						received = slides;
					},
					getConfig: () => CONFIG,
					getAllowExternalContent: () => allow,
					createSession: sessionOn(doc),
				});
				await collab.start(CONFIG);
				await vi.waitFor(() => expect(received[0]?.elements).toHaveLength(1));
				expect((received[0]?.elements[0] as { imageData?: string } | undefined)?.imageData).toBe(
					expected,
				);
				collab.stop();
			});
			doc.destroy();
		},
	);
});

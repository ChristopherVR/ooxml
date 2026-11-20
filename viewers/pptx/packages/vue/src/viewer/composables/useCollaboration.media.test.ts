/**
 * The room carries the deck's media, and linked content follows this viewer's own setting:
 * the composable hands the room the URLs the viewer plays its video from, and reads the
 * room's slides with the viewer's remote content setting.
 */
import type { CollaborationConfig, ExternalCollaborationSession } from 'ooxml-ui/pptx';
import { reconcileSlidesInYDoc } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { effectScope, ref } from 'vue';
import * as Y from 'yjs';

import { useCollaboration } from './useCollaboration';

const factories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
};
const VIDEO_PATH = 'ppt/media/media1.mp4';
const PLAYED_FROM = 'blob:http://localhost/played-video';
const LINKED = 'https://linked.test/picture.png';

const deck = (...elements: unknown[]): PptxSlide[] =>
	[{ id: 's1', rId: 'rId2', slideNumber: 1, elements }] as unknown as PptxSlide[];

const cleanup: (() => void)[] = [];
afterEach(() => {
	for (const dispose of cleanup.splice(0).reverse()) {
		dispose();
	}
});

function hostSession(doc: Y.Doc): ExternalCollaborationSession {
	let local: Record<string, unknown> | null = { user: 'host' };
	return {
		doc,
		awareness: {
			clientID: doc.clientID,
			getLocalState: () => local,
			setLocalState: (state) => {
				local = state;
			},
			setLocalStateField: (field, value) => {
				local = { ...local, [field]: value };
			},
			getStates: () => new Map(local ? [[doc.clientID, local]] : []),
			on: () => {},
			off: () => {},
		},
		getSnapshot: () => ({ status: 'connected', synced: true }),
		subscribe: () => () => {},
	};
}

async function share(initial: PptxSlide[], allowExternalContent: boolean, doc: Y.Doc) {
	const scope = effectScope();
	cleanup.push(() => scope.stop());
	const slides = ref(initial);
	const collab = scope.run(() =>
		useCollaboration({
			slides,
			onRemoteSlides: (next) => {
				slides.value = next;
			},
			getMediaDataUrls: () => new Map([[VIDEO_PATH, PLAYED_FROM]]),
			allowExternalContent: () => allowExternalContent,
		}),
	)!;
	const config: CollaborationConfig = {
		roomId: 'media-room',
		serverUrl: '',
		userName: 'Ada',
		externalSession: hostSession(doc),
	};
	await collab.start(config);
	return slides;
}

describe('collaboration media in the Vue binding', () => {
	it('shares an embedded video as a media token, never the URL it plays from', async () => {
		const doc = new Y.Doc();
		cleanup.push(() => doc.destroy());
		const video = { id: 'vid', type: 'media', mediaType: 'video', mediaPath: VIDEO_PATH };
		await share(deck(video), false, doc);
		await vi.waitFor(() => expect(doc.getArray('pptx:slides')).toHaveLength(1));
		const assets = doc.getMap('pptx:assets').toJSON() as Record<string, string>;
		expect(assets['vid:mediaData']).toMatch(/^ooxml-media:[0-9a-f]{32}$/u);
		expect(JSON.stringify(doc.toJSON())).not.toContain('blob:');
	});

	it.each([
		[false, undefined],
		[true, LINKED],
	])(
		'with external content allowed: %s, shows a linked picture from the room as %s',
		async (allow, expected) => {
			const doc = new Y.Doc();
			cleanup.push(() => doc.destroy());
			reconcileSlidesInYDoc(
				deck({ id: 'pic', type: 'picture', imagePath: LINKED, imageData: LINKED }),
				doc,
				factories,
				'remote',
			);
			const slides = await share(deck(), allow, doc);
			await vi.waitFor(() => expect(slides.value[0]?.elements).toHaveLength(1));
			expect((slides.value[0]?.elements[0] as { imageData?: string } | undefined)?.imageData).toBe(
				expected,
			);
		},
	);
});

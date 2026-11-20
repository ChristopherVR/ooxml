/**
 * The room carries the deck's media, and linked content follows this viewer's own setting:
 * the service hands the room the URLs the viewer plays its video from, and reads the room's
 * slides with the viewer's remote content setting.
 */
import { DestroyRef, Injector } from '@angular/core';
import type { CollaborationConfig, ExternalCollaborationSession } from 'ooxml-ui/pptx';
import { reconcileSlidesInYDoc } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import { CollaborationService } from './collaboration.service';

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

const docs: Y.Doc[] = [];
afterEach(() => {
	for (const doc of docs.splice(0)) {
		doc.destroy();
	}
});

function room(): { doc: Y.Doc; session: ExternalCollaborationSession } {
	const doc = new Y.Doc();
	docs.push(doc);
	let state: Record<string, unknown> | null = {};
	const session: ExternalCollaborationSession = {
		doc,
		awareness: {
			clientID: doc.clientID,
			getLocalState: () => state,
			setLocalState: (next) => {
				state = next;
			},
			setLocalStateField: (key, value) => {
				state = { ...state, [key]: value };
			},
			getStates: () => new Map(state ? [[doc.clientID, state]] : []),
			on: () => {},
			off: () => {},
		},
		getSnapshot: () => ({ status: 'connected', synced: true }),
		subscribe: () => () => {},
	};
	return { doc, session };
}

function service(): CollaborationService {
	return Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: { onDestroy: () => () => undefined } },
			{ provide: CollaborationService, useClass: CollaborationService },
		],
	}).get(CollaborationService);
}

const config = (session: ExternalCollaborationSession): CollaborationConfig => ({
	roomId: 'media-room',
	serverUrl: 'ws://unused.invalid',
	userName: 'Ada',
	externalSession: session,
});

describe('collaboration media in the Angular binding', () => {
	it('shares an embedded video as a media token, never the URL it plays from', async () => {
		const { doc, session } = room();
		const collab = service();
		const slides = deck({ id: 'vid', type: 'media', mediaType: 'video', mediaPath: VIDEO_PATH });
		await collab.connect(
			{ ...config(session), sessionIntent: 'create' },
			{ getMediaDataUrls: () => new Map([[VIDEO_PATH, PLAYED_FROM]]) },
		);
		collab.broadcastSlides(slides);
		await vi.waitFor(() => expect(doc.getArray('pptx:slides')).toHaveLength(1));
		const assets = doc.getMap('pptx:assets').toJSON() as Record<string, string>;
		expect(assets['vid:mediaData']).toMatch(/^ooxml-media:[0-9a-f]{32}$/u);
		expect(JSON.stringify(doc.toJSON())).not.toContain('blob:');
		collab.disconnect();
	});

	it.each([
		[false, undefined],
		[true, LINKED],
	])(
		'with external content allowed: %s, shows a linked picture from the room as %s',
		async (allow, expected) => {
			const { doc, session } = room();
			reconcileSlidesInYDoc(
				deck({ id: 'pic', type: 'picture', imagePath: LINKED, imageData: LINKED }),
				doc,
				factories,
				'remote',
			);
			const collab = service();
			let received: PptxSlide[] = [];
			await collab.connect(config(session), {
				onRemoteSlides: (slides) => {
					received = slides;
				},
				allowExternalContent: () => allow,
			});
			await vi.waitFor(() => expect(received[0]?.elements).toHaveLength(1));
			expect((received[0]?.elements[0] as { imageData?: string } | undefined)?.imageData).toBe(
				expected,
			);
			collab.disconnect();
		},
	);
});

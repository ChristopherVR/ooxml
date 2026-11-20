// @vitest-environment happy-dom
/**
 * The room carries the deck's media, and linked content follows this viewer's own setting:
 * the hook hands the room the URLs the viewer plays its video from, and reads the room's
 * slides with the viewer's remote content setting.
 */
import type { YDocLike, YjsFactories } from 'ooxml-ui/pptx';
import { reconcileSlidesInYDoc } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import React, { act, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import { useYjsDocumentSync } from './useYjsDocumentSync';

const factories: YjsFactories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
};
const VIDEO_PATH = 'ppt/media/media1.mp4';
const PLAYED_FROM = 'blob:http://localhost/played-video';
const LINKED = 'https://linked.test/picture.png';

const deck = (...elements: unknown[]): PptxSlide[] =>
	[{ id: 's1', rId: 'rId2', slideNumber: 1, elements }] as unknown as PptxSlide[];

let root: Root;
let container: HTMLDivElement;
let doc: Y.Doc;
let current: PptxSlide[];

function Probe({
	initial,
	allowExternalContent,
}: {
	initial: PptxSlide[];
	allowExternalContent: boolean;
}): null {
	const [slides, setSlides] = useState(initial);
	current = slides;
	useYjsDocumentSync({
		doc,
		slides,
		setSlides,
		isConnected: true,
		isSynced: true,
		templateElementsBySlideId: {},
		mediaDataUrls: new Map([[VIDEO_PATH, PLAYED_FROM]]),
		allowExternalContent,
	});
	return null;
}

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	doc = new Y.Doc();
});

afterEach(() => {
	act(() => root.unmount());
	doc.destroy();
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('collaboration media in the React binding', () => {
	it('shares an embedded video as a media token, never the URL it plays from', async () => {
		const video = { id: 'vid', type: 'media', mediaType: 'video', mediaPath: VIDEO_PATH };
		await act(async () =>
			root.render(<Probe initial={deck(video)} allowExternalContent={false} />),
		);
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
			reconcileSlidesInYDoc(
				deck({ id: 'pic', type: 'picture', imagePath: LINKED, imageData: LINKED }),
				doc as unknown as YDocLike,
				factories,
				'remote',
			);
			await act(async () => root.render(<Probe initial={deck()} allowExternalContent={allow} />));
			const picture = current[0]?.elements[0] as { imageData?: string } | undefined;
			expect(picture?.imageData).toBe(expected);
		},
	);
});

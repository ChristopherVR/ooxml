// @vitest-environment happy-dom
/**
 * Joining a room whose slides replace the open deck is a presentation load (it
 * drops the assistant's staged proposals); a peer's later edit, or a room this
 * editor seeded, is not. Covers the built-in y-websocket path and a host session.
 */
import type { CollaborationConfig, YDocLike, YjsFactories } from 'ooxml-ui/pptx';
import { reconcileSlidesInYDoc } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import React, { act, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import { useYjsDocumentSync } from './useYjsDocumentSync';

const slide = (id: string): PptxSlide => ({ id, rId: id, slideNumber: 1, elements: [] });
const factories: YjsFactories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
};

let root: Root;
let container: HTMLDivElement;
let doc: Y.Doc;
let current: PptxSlide[];

/** A peer writes `id` as the room's only slide. */
function peerWrites(id: string): void {
	reconcileSlidesInYDoc([slide(id)], doc as unknown as YDocLike, factories, 'remote');
}

function Probe({
	hostSession,
	intent,
	onRoomReplacedDeck,
}: {
	hostSession: boolean;
	intent?: 'join';
	onRoomReplacedDeck: () => void;
}): null {
	const [slides, setSlides] = useState([slide('local')]);
	current = slides;
	const config = useMemo<CollaborationConfig | undefined>(
		() =>
			intent
				? ({ sessionIntent: intent } as CollaborationConfig)
				: hostSession
					? ({
							sessionIntent: 'join',
							externalSession: {
								doc,
								getSnapshot: () => ({ status: 'connected', synced: true }),
								subscribe: () => () => {},
							},
						} as CollaborationConfig)
					: undefined,
		[hostSession, intent],
	);
	useYjsDocumentSync({
		doc,
		slides,
		setSlides,
		isConnected: true,
		isSynced: true,
		config,
		loadOrigin: 'bootstrap',
		templateElementsBySlideId: {},
		onRoomReplacedDeck,
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

describe.each([
	['the built-in sync', false],
	['a host session', true],
])('room joins through %s', (_name, hostSession) => {
	it('reports a join that replaces the open deck once, and not a peer edit after it', async () => {
		peerWrites('room');
		const replaced = vi.fn();
		await act(async () =>
			root.render(<Probe hostSession={hostSession} onRoomReplacedDeck={replaced} />),
		);
		expect(current.map((item) => item.id)).toStrictEqual(['room']);
		expect(replaced).toHaveBeenCalledOnce();

		await act(async () => peerWrites('peer-edit'));

		expect(current.map((item) => item.id)).toStrictEqual(['peer-edit']);
		expect(replaced).toHaveBeenCalledOnce();
	});
});

describe('a session that joins a room', () => {
	it('never seeds the empty room with its placeholder deck, and adopts the sharer deck', async () => {
		const replaced = vi.fn();
		await act(async () =>
			root.render(<Probe hostSession={false} intent='join' onRoomReplacedDeck={replaced} />),
		);
		await act(async () => {
			await new Promise<void>((resolve) => {
				setTimeout(resolve, 20);
			});
		});
		expect(doc.getArray('pptx:slides')).toHaveLength(0);

		await act(async () => peerWrites('sharer'));

		expect(current.map((item) => item.id)).toStrictEqual(['sharer']);
		expect(replaced).toHaveBeenCalledOnce();
	});
});

describe('a room this editor seeded', () => {
	it('does not report a peer edit as a join', async () => {
		const replaced = vi.fn();
		await act(async () => root.render(<Probe hostSession={false} onRoomReplacedDeck={replaced} />));
		await vi.waitFor(() => expect(doc.getArray('pptx:slides')).toHaveLength(1));

		await act(async () => peerWrites('peer-edit'));

		expect(current.map((item) => item.id)).toStrictEqual(['peer-edit']);
		expect(replaced).not.toHaveBeenCalled();
	});
});

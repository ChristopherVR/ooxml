/**
 * The built-in (y-websocket) sync engine over a real Y.Doc: which room adoptions replace the open
 * deck, so the viewer reports them as a presentation load, and which are a peer's ordinary edit.
 */
import type { YjsFactories } from 'ooxml-ui/pptx';
import { observeYDocSlides, reconcileSlidesInYDoc } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import { SlideSyncEngine } from './collaboration-slide-sync';

const factories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
} as unknown as YjsFactories;
const slide = (id: string): PptxSlide => ({ id, elements: [] }) as unknown as PptxSlide;

function room(...existing: string[]) {
	const doc = new Y.Doc();
	reconcileSlidesInYDoc(existing.map(slide), doc as never, factories, 'peer');
	const sync = new SlideSyncEngine();
	const replaced = vi.fn();
	sync.bind({
		ydoc: doc as never,
		factories,
		onRemoteSlides: () => undefined,
		onRoomReplacedDeck: replaced,
		scheduleWriteBack: () => undefined,
	});
	observeYDocSlides(doc as never, (_events, transaction) => sync.onRemoteChange(transaction));
	const peerWrites = (...ids: string[]) =>
		reconcileSlidesInYDoc(ids.map(slide), doc as never, factories, 'peer');
	return { doc, sync, replaced, peerWrites };
}

describe('room slides that replace the open deck', () => {
	it('reports a join once, and not the peer edits that follow it', () => {
		const { replaced, peerWrites } = room();
		peerWrites('room-1');
		expect(replaced).toHaveBeenCalledOnce();
		peerWrites('room-1', 'room-2');
		expect(replaced).toHaveBeenCalledOnce();
	});

	it('reports nothing once this editor seeded the room with its own deck', () => {
		const { sync, replaced, peerWrites } = room();
		sync.gate.open();
		sync.broadcast([slide('mine')]);
		peerWrites('mine', 'peer');
		expect(replaced).not.toHaveBeenCalled();
	});

	it('reports a room whose slides replace the bootstrap deck that loaded after joining', () => {
		const { sync, replaced } = room('room-1');
		expect(sync.adoptDocAfterLoad('bootstrap')).toBeTruthy();
		expect(replaced).toHaveBeenCalledOnce();
	});
});

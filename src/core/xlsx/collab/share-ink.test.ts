import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { createEditSession } from '../edit/session';
import type { EditSession } from '../edit/types';
import type { InkObject, Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { type WorkbookCollabHost, bindWorkbookSession } from './bind';
import { inkSpec } from './part-specs';

interface Peer {
	doc: Y.Doc;
	edit: EditSession;
	open: () => void;
}

function peer(workbook: Workbook, writable = true): Peer {
	const doc = new Y.Doc();
	const ready = new Set<(value: undefined) => void>();
	const state = { writable };
	const host = {
		doc,
		synced: writable,
		canWrite: () => state.writable,
		on: (event: string, listener: (value: undefined) => void) => {
			if (event !== 'ready') return () => undefined;
			ready.add(listener);
			return () => ready.delete(listener);
		},
		updatePresence: () => undefined,
		peers: () => [],
	} as unknown as WorkbookCollabHost;
	const edit = createEditSession(workbook);
	bindWorkbookSession(host, edit);
	return {
		doc,
		edit,
		open: () => {
			state.writable = true;
			for (const listener of ready) listener(undefined);
		},
	};
}

function sync(...peers: Peer[]): void {
	for (const a of peers)
		for (const b of peers)
			if (a !== b) Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(a.doc, Y.encodeStateVector(b.doc)));
}

// Excel 16's own save, with ink pasted from PowerPoint and its fallback picture.
const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-ink-fallback.xlsx')),
	);

async function pair(): Promise<[Peer, Peer]> {
	const a = peer(await loadXlsx(fixture()));
	const b = peer(await loadXlsx(fixture()), false);
	sync(a, b);
	b.open();
	sync(a, b);
	return [a, b];
}

const inksOf = (p: Peer): InkObject[] =>
	p.edit.workbook.sheets[0]!.drawings.filter((d): d is InkObject => d.kind === 'ink');
const anchor = (row: number) => ({ from: { row, col: 1, colOffset: 0, rowOffset: 0 } });
const stroke = (color: string) => ({
	path: 'M 0 0 L 20 10 L 40 0',
	color,
	width: 2,
	opacity: 1,
});

describe('sharing Excel ink', () => {
	it('sends a stroke drawn by one peer to the other, keeping each peer its own part', async () => {
		const [a, b] = await pair();
		const before = inksOf(b)[0];
		expect(before?.partName).toBe('xl/ink/ink1.xml');
		a.edit.addInk(0, anchor(20), { strokes: [stroke('#FF0000')], width: 40, height: 10 });
		sync(a, b);
		expect(inksOf(b)).toHaveLength(2);
		expect(inksOf(b)[1]?.strokes[0]?.color).toBe('#FF0000');
		// The ink both opened keeps the part and anchor XML of b's own package.
		expect(inksOf(b)[0]?.partName).toBe('xl/ink/ink1.xml');
		expect(inksOf(b)[0]?.sourceXml).toBe(before?.sourceXml);
	});

	it('keeps ink two peers draw at the same time, and takes the last edit of one drawing', async () => {
		const [a, b] = await pair();
		a.edit.addInk(0, anchor(20), { strokes: [stroke('#FF0000')], width: 40, height: 10 });
		b.edit.addInk(0, anchor(30), { strokes: [stroke('#0000FF')], width: 40, height: 10 });
		sync(a, b);
		for (const p of [a, b])
			expect(inksOf(p).map((ink) => ink.strokes[0]?.color)).toEqual(
				expect.arrayContaining(['#FF0000', '#0000FF']),
			);
		const index = a.edit.workbook.sheets[0]!.drawings.findIndex(
			(d) => d.kind === 'ink' && d.strokes[0]?.color === '#FF0000',
		);
		a.edit.setInk(0, index, anchor(20), {
			strokes: [stroke('#FF0000'), stroke('#00FF00')],
			width: 40,
			height: 10,
		});
		sync(a, b);
		expect(inksOf(b).find((ink) => ink.strokes[0]?.color === '#FF0000')?.strokes).toHaveLength(2);
		// Erasing every stroke deletes the drawing for everyone.
		a.edit.setInk(0, index, anchor(20), { strokes: [], width: 40, height: 10 });
		sync(a, b);
		expect(inksOf(b).some((ink) => ink.strokes[0]?.color === '#FF0000')).toBe(false);
	});

	it('leaves a peer its own ink in a room written by a build that kept ink local', async () => {
		const a = peer(await loadXlsx(fixture()));
		const old = new Y.Doc();
		Y.applyUpdate(old, Y.encodeStateAsUpdate(a.doc));
		old.getMap('xlsx:meta').delete('ink');
		old.getMap('xlsx:ink').clear();
		const b = peer(await loadXlsx(fixture()), false);
		Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(old));
		b.open();
		expect(inksOf(b)).toHaveLength(1);
	});

	it('refuses malformed or partial ink from a peer and never shares truncated ink', () => {
		const ok = {
			kind: 'ink',
			anchor: anchor(1),
			strokes: [stroke('#000000')],
			width: 4,
			height: 4,
		};
		expect(inkSpec.decode(ok)).toBeDefined();
		expect(
			inkSpec.decode({ ...ok, strokes: [{ ...stroke('red'), color: 'red' }] }),
		).toBeUndefined();
		expect(inkSpec.decode({ ...ok, width: 0 })).toBeUndefined();
		expect(inkSpec.decode({ ...ok, sourceXml: '<x/>' })).toBeUndefined();
		expect(inkSpec.decode({ ...ok, truncated: true })).toBeUndefined();
		expect(inkSpec.encode({ ...(ok as InkObject), truncated: true })).toBeUndefined();
		expect(
			inkSpec.encode({ ...(ok as InkObject), partName: 'xl/ink/ink9.xml', sourceXml: '<a/>' }),
		).not.toHaveProperty('partName');
	});
});

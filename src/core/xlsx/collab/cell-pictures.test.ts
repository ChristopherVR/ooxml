import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { encodePng } from '../../drawingml/raster-encode';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import type { EditSession } from '../edit/types';
import { cellView } from '../layout/cell-view';
import type { Workbook } from '../model';
import { loadXlsx } from '../read/load';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/save';
import { type WorkbookCollabHost, bindWorkbookSession } from './bind';
import { MAX_SHARED_CELL_PICTURE_BYTES, PictureBridge } from './cell-pictures';
import { MEDIA, sharedWorkbook } from './schema';

interface Peer {
	doc: Y.Doc;
	edit: EditSession;
	undo: () => boolean;
}

function peer(workbook: Workbook, from?: Peer): Peer {
	const doc = new Y.Doc();
	if (from) Y.applyUpdate(doc, Y.encodeStateAsUpdate(from.doc));
	const host = {
		doc,
		synced: true,
		canWrite: () => true,
		on: () => () => undefined,
		updatePresence: () => undefined,
		peers: () => [],
	} as unknown as WorkbookCollabHost;
	const edit = createEditSession(workbook);
	const binding = bindWorkbookSession(host, edit);
	return { doc, edit, undo: binding.undo };
}

function sync(...peers: Peer[]): void {
	for (const a of peers)
		for (const b of peers)
			if (a !== b) Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(a.doc, Y.encodeStateVector(b.doc)));
}

const PNG = encodePng(2, 1, Uint8Array.of(200, 50, 50, 50, 50, 200));
const at = (row: number, col: number) => ({ row, col });
const pictureAt = (p: Peer, row: number, col: number) =>
	cellView(p.edit.workbook, 0, row, col).picture;

describe('sharing in-cell pictures', () => {
	it('shares placed pictures, alt text and removal, storing the bytes once', () => {
		const a = peer(createWorkbook());
		const b = peer(createWorkbook(), a);
		expect(a.edit.addCellPicture(0, at(0, 0), PNG, 'Logo')).toBe(true);
		expect(a.edit.addCellPicture(0, at(1, 0), PNG, 'Again')).toBe(true);
		sync(a, b);
		expect(pictureAt(b, 0, 0)?.bytes).toEqual(PNG);
		expect(pictureAt(b, 0, 0)?.altText).toBe('Logo');
		expect(pictureAt(b, 1, 0)?.altText).toBe('Again');
		expect(b.doc.getMap(MEDIA).size).toBe(1);
		const parts = [0, 1].map((row) => getCell(b.edit.workbook.sheets[0]!, row, 0)?.inCellPicture);
		expect(parts[0]?.partName).toBe(parts[1]?.partName);

		b.edit.setCellPictureAltText(0, at(0, 0), 'Company logo');
		sync(a, b);
		expect(pictureAt(a, 0, 0)?.altText).toBe('Company logo');

		a.edit.clearRange(0, { start: at(1, 0), end: at(1, 0) }, 'all');
		sync(a, b);
		expect(pictureAt(b, 1, 0)).toBeUndefined();
	});

	it('keeps the bytes in the room when the picture is undone', () => {
		const a = peer(createWorkbook());
		const b = peer(createWorkbook(), a);
		a.edit.addCellPicture(0, at(0, 0), PNG);
		sync(a, b);
		b.edit.addCellPicture(0, at(2, 2), PNG);
		a.undo();
		sync(a, b);
		expect(pictureAt(b, 0, 0)).toBeUndefined();
		expect(pictureAt(a, 2, 2)?.bytes).toEqual(PNG);
		expect(a.doc.getMap(MEDIA).size).toBe(1);
	});

	it('shares pictures loaded from a file and saves them on the peer', async () => {
		const seed = createWorkbook();
		createEditSession(seed).addCellPicture(0, at(0, 0), PNG, 'Loaded');
		const loaded = await loadXlsx(await saveXlsx(seed));
		expect(getCell(loaded.sheets[0]!, 0, 0)?.valueMetadata).toBeDefined();
		const a = peer(loaded);
		const b = peer(createWorkbook(), a);
		expect(pictureAt(b, 0, 0)?.altText).toBe('Loaded');
		const reopened = await loadXlsx(await saveXlsx(b.edit.workbook));
		expect(cellView(reopened, 0, 0, 0).picture?.bytes).toEqual(PNG);
	});

	it('refuses forged, mistyped and oversized bytes', () => {
		const doc = new Y.Doc();
		const shared = sharedWorkbook(doc);
		const workbook = createWorkbook();
		const bridge = new PictureBridge(workbook, shared.media, false);
		const hash = 'ab'.repeat(32);
		shared.media.set(hash, { t: 'image/png', b: PNG });
		expect(bridge.pictureOf({ h: hash })).toBeUndefined();
		expect(bridge.pictureOf({ h: 'nothex' })).toBeUndefined();
		expect(bridge.pictureOf({ h: hash, a: 'x'.repeat(2000) })).toBeUndefined();
		const writer = new PictureBridge(workbook, shared.media, true);
		workbook.source = { parts: new Map([['xl/media/image1.png', PNG]]) };
		const cell = {
			value: { error: '#VALUE!' as const },
			inCellPicture: { partName: 'xl/media/image1.png', altText: '' },
		};
		const ref = writer.refOf(cell);
		expect(ref?.h).toMatch(/^[0-9a-f]{64}$/);
		shared.media.set(ref!.h, { t: 'image/gif', b: PNG });
		const reader = new PictureBridge(createWorkbook(), shared.media, false);
		expect(reader.pictureOf(ref)).toBeUndefined();
		const big = new Uint8Array(MAX_SHARED_CELL_PICTURE_BYTES + 1);
		big.set(PNG);
		workbook.source.parts.set('xl/media/image1.png', big);
		expect(new PictureBridge(workbook, shared.media, true).refOf(cell)).toBeUndefined();
	});
});

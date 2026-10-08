import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import type { EditSession, WorkbookChange } from '../edit/types';
import type { Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { type WorkbookBinding, type WorkbookCollabHost, bindWorkbookSession } from './bind';
import type { XlsxPresence } from './presence';

interface Peer {
	doc: Y.Doc;
	edit: EditSession;
	binding: WorkbookBinding;
	open: () => void;
	block: () => void;
	presence: XlsxPresence;
	changes: WorkbookChange[];
}

/** A peer on a fake session: writable at once (the room creator) or after `open()`. */
function peer(workbook: Workbook = createWorkbook(), writable = true): Peer {
	const doc = new Y.Doc();
	const ready = new Set<(value: undefined) => void>();
	const state = { writable, presence: {} as XlsxPresence };
	const host = {
		doc,
		synced: writable,
		canWrite: () => state.writable,
		on: (event: string, listener: (value: undefined) => void) => {
			if (event !== 'ready') return () => undefined;
			ready.add(listener);
			return () => ready.delete(listener);
		},
		updatePresence: (patch: XlsxPresence) => Object.assign(state.presence, patch),
		peers: () => [],
	} as unknown as WorkbookCollabHost;
	const edit = createEditSession(workbook);
	const changes: WorkbookChange[] = [];
	edit.onChange((change) => changes.push(change));
	const binding = bindWorkbookSession(host, edit);
	const open = (): void => {
		state.writable = true;
		for (const listener of ready) listener(undefined);
	};
	const block = (): void => {
		state.writable = false;
	};
	return { doc, edit, binding, open, block, presence: state.presence, changes };
}

/** Exchanges every missing update between the documents, as a provider would. */
function sync(...peers: Peer[]): void {
	for (const a of peers)
		for (const b of peers)
			if (a !== b) Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(a.doc, Y.encodeStateVector(b.doc)));
}

/** A second peer joining an existing room. */
function join(host: Peer): Peer {
	const guest = peer(createWorkbook(), false);
	sync(host, guest);
	guest.open();
	return guest;
}

const value = (p: Peer, ref: string, sheet = 0) => {
	const at = /^([A-Z]+)(\d+)$/.exec(ref);
	const col = (at?.[1] ?? 'A').charCodeAt(0) - 65;
	const row = Number(at?.[2] ?? 1) - 1;
	const target = p.edit.workbook.sheets[sheet];
	return target ? (getCell(target, row, col)?.value ?? null) : undefined;
};
const names = (p: Peer) => p.edit.workbook.sheets.map((s) => s.name);

describe('bindWorkbookSession', () => {
	it('seeds an empty room and a joining peer adopts it', () => {
		const a = peer();
		a.edit.setCellInput(0, 0, 0, 'hello');
		a.edit.addSheet('Data');
		const b = join(a);
		expect(names(b)).toEqual(['Sheet1', 'Data']);
		expect(value(b, 'A1')).toBe('hello');
		expect(b.edit.canUndo()).toBe(false);
		expect(b.binding.canUndo()).toBe(false);
	});

	it('keeps concurrent edits to different cells', () => {
		const a = peer();
		const b = join(a);
		a.edit.setCellInput(0, 0, 0, '1');
		b.edit.setCellInput(0, 1, 1, 'x');
		sync(a, b);
		for (const p of [a, b]) {
			expect(value(p, 'A1')).toBe(1);
			expect(value(p, 'B2')).toBe('x');
		}
		const remote = b.changes.at(-1);
		expect(remote).toMatchObject({ kind: 'remote', external: true, sheet: 0 });
	});

	it('converges on one value when two peers edit the same cell', () => {
		const a = peer();
		const b = join(a);
		a.edit.setCellInput(0, 0, 0, 'from a');
		b.edit.setCellInput(0, 0, 0, 'from b');
		sync(a, b);
		expect(value(a, 'A1')).toBe(value(b, 'A1'));
		expect(['from a', 'from b']).toContain(value(a, 'A1'));
	});

	it('shares formulas with their values and recalculates remote inputs', () => {
		const a = peer();
		const b = join(a);
		a.edit.setCellInput(0, 0, 0, '2');
		a.edit.setCellInput(0, 0, 1, '=A1*3');
		sync(a, b);
		expect(getCell(b.edit.workbook.sheets[0]!, 0, 1)).toMatchObject({ formula: 'A1*3', value: 6 });
		b.edit.setCellInput(0, 0, 0, '5');
		sync(a, b);
		expect(value(a, 'B1')).toBe(15);
	});

	it('undoes only local edits and shares the undo', () => {
		const a = peer();
		const b = join(a);
		a.edit.setCellInput(0, 0, 0, 'mine');
		a.edit.setCellInput(0, 2, 0, 'mine too');
		b.edit.setCellInput(0, 1, 1, 'theirs');
		sync(a, b);
		expect(a.binding.undo()).toBe(true);
		expect(value(a, 'A3')).toBe(null);
		expect(value(a, 'A1')).toBe('mine');
		expect(value(a, 'B2')).toBe('theirs');
		expect(a.changes.at(-1)).toMatchObject({ kind: 'undo', external: true });
		sync(a, b);
		expect(value(b, 'A3')).toBe(null);
		expect(a.binding.redo()).toBe(true);
		sync(a, b);
		expect(value(b, 'A3')).toBe('mine too');
		// B's stack holds only its own edit.
		expect(b.binding.undo()).toBe(true);
		expect(b.binding.undo()).toBe(false);
		expect(value(b, 'B2')).toBe(null);
		expect(value(b, 'A1')).toBe('mine');
	});

	it('syncs sheet add, rename, move and delete', () => {
		const a = peer();
		const b = join(a);
		a.edit.addSheet('Two');
		a.edit.addSheet('Three');
		sync(a, b);
		expect(names(b)).toEqual(['Sheet1', 'Two', 'Three']);
		b.edit.setCellInput(1, 0, 0, 'on two');
		b.edit.renameSheet(1, 'Second');
		sync(a, b);
		expect(names(a)).toEqual(['Sheet1', 'Second', 'Three']);
		expect(value(a, 'A1', 1)).toBe('on two');
		a.edit.moveSheet(2, 0);
		sync(a, b);
		expect(names(b)).toEqual(['Three', 'Sheet1', 'Second']);
		b.edit.deleteSheet(0);
		sync(a, b);
		expect(names(a)).toEqual(['Sheet1', 'Second']);
		expect(a.edit.workbook.sheets[a.edit.workbook.activeSheet]).toBeDefined();
	});

	it('makes sheets added concurrently under one name unique on every peer', () => {
		const a = peer();
		const b = join(a);
		a.edit.addSheet('Sheet2');
		b.edit.addSheet('Sheet2');
		a.edit.setCellInput(1, 0, 0, 'a');
		b.edit.setCellInput(1, 0, 0, 'b');
		sync(a, b);
		expect(names(a)).toEqual(names(b));
		expect(new Set(names(a).map((n) => n.toLowerCase())).size).toBe(3);
		const ids = a.edit.workbook.sheets.map((s) => s.sheetId);
		expect(new Set(ids).size).toBe(3);
		expect(b.edit.workbook.sheets.map((s) => s.sheetId)).toEqual(ids);
		const cells = [1, 2].map((i) => value(a, 'A1', i)).sort();
		expect(cells).toEqual(['a', 'b']);
	});

	it('concurrent moves of different sheets both apply', () => {
		const a = peer(createWorkbook({ sheets: ['S1', 'S2', 'S3', 'S4'] }));
		const b = join(a);
		a.edit.moveSheet(0, 3);
		b.edit.moveSheet(1, 0);
		sync(a, b);
		expect(names(a)).toEqual(names(b));
		expect(names(a).length).toBe(4);
	});

	it('a deleted sheet wins over a concurrent edit on it', () => {
		const a = peer(createWorkbook({ sheets: ['S1', 'S2'] }));
		const b = join(a);
		a.edit.deleteSheet(1);
		b.edit.setCellInput(1, 4, 4, 'lost');
		sync(a, b);
		expect(names(a)).toEqual(['S1']);
		expect(names(b)).toEqual(['S1']);
	});

	it('syncs row heights, column widths and merges', () => {
		const a = peer();
		const b = join(a);
		a.edit.setRowHeight(0, [2], 30);
		b.edit.setColumnWidth(0, [3], 20);
		a.edit.setCellInput(0, 5, 0, 'merged');
		a.edit.merge(0, { start: { row: 5, col: 0 }, end: { row: 6, col: 2 } }, 'merge');
		sync(a, b);
		for (const p of [a, b]) {
			const sheet = p.edit.workbook.sheets[0]!;
			expect(sheet.rowInfo.get(2)?.height).toBe(30);
			expect(sheet.columns.find((c) => c.min <= 3 && c.max >= 3)?.width).toBe(20);
			expect(sheet.merges).toHaveLength(1);
		}
		expect(b.changes.some((c) => c.external && c.structural)).toBe(true);
	});

	it('resolves concurrent column splits to the most specific width', () => {
		const a = peer();
		const b = join(a);
		a.edit.setColumnWidth(0, [2], 25);
		b.edit.setColumnWidth(0, [4], 12);
		sync(a, b);
		for (const p of [a, b]) {
			const cols = p.edit.workbook.sheets[0]!.columns;
			expect(cols.find((c) => c.min <= 2 && c.max >= 2)?.width).toBe(25);
			expect(cols.find((c) => c.min <= 4 && c.max >= 4)?.width).toBe(12);
		}
	});

	it('shares formats by content, not by local style id', () => {
		const a = peer();
		const b = join(a);
		b.edit.applyStyle(0, [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }], {
			font: { italic: true },
		});
		a.edit.applyStyle(0, [{ start: { row: 1, col: 0 }, end: { row: 1, col: 0 } }], {
			font: { bold: true },
		});
		sync(a, b);
		for (const p of [a, b]) {
			const wb = p.edit.workbook;
			const sheet = wb.sheets[0]!;
			expect(wb.styles[getCell(sheet, 0, 0)?.styleId ?? 0]?.font.italic).toBe(true);
			expect(wb.styles[getCell(sheet, 1, 0)?.styleId ?? 0]?.font.bold).toBe(true);
		}
	});

	it('catches up from a backlog of updates after a reconnect', () => {
		const a = peer();
		const b = join(a);
		const backlog: Uint8Array[] = [];
		a.doc.on('update', (update: Uint8Array) => backlog.push(update));
		for (let i = 0; i < 20; i++) a.edit.setCellInput(0, i, 0, String(i));
		a.edit.addSheet('Later');
		a.edit.renameSheet(0, 'First');
		b.edit.setCellInput(0, 0, 3, 'offline');
		for (const update of backlog) Y.applyUpdate(b.doc, update);
		sync(a, b);
		expect(names(b)).toEqual(['First', 'Later']);
		expect(value(b, 'A20')).toBe(19);
		expect(value(a, 'D1')).toBe('offline');
		expect(b.edit.canUndo()).toBe(true); // Only its own offline edit.
	});

	it('publishes edits made while writes were blocked once they are allowed', () => {
		const a = peer();
		const b = join(a);
		b.block();
		b.edit.setCellInput(0, 3, 3, 'late');
		sync(a, b);
		expect(value(a, 'D4')).toBe(null);
		b.open();
		sync(a, b);
		expect(value(a, 'D4')).toBe('late');
	});

	it('publishes the selection with the sheet key and resolves it back', () => {
		const a = peer();
		a.binding.setSelection(0, { start: { row: 1, col: 1 }, end: { row: 3, col: 2 } });
		expect(a.presence.range).toBe('B2:C4');
		expect(a.binding.sheetIndex(a.presence.sheet ?? '')).toBe(0);
		a.edit.addSheet('Before', 0);
		expect(a.binding.sheetIndex(a.presence.sheet ?? '')).toBe(1);
	});

	it('stops syncing after dispose', () => {
		const a = peer();
		const b = join(a);
		b.binding.dispose();
		a.edit.setCellInput(0, 0, 0, 'after');
		sync(a, b);
		expect(value(b, 'A1')).toBe(null);
	});
});

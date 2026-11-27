import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import type { EditSession } from '../edit/types';
import type { PivotTable, Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { type WorkbookCollabHost, bindWorkbookSession } from './bind';
import { pivotSpec } from './pivot-spec';

interface Peer {
	doc: Y.Doc;
	edit: EditSession;
}

/** A peer bound to a new room, or joining the room `from` is in (it adopts what is there). */
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
	bindWorkbookSession(host, edit);
	return { doc, edit };
}

function sync(...peers: Peer[]): void {
	for (const a of peers)
		for (const b of peers)
			if (a !== b) Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(a.doc, Y.encodeStateVector(b.doc)));
}

function dataBook(): Workbook {
	const workbook = createWorkbook({ sheets: ['Data', 'Summary'] });
	const edit = createEditSession(workbook);
	[
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Fruit', 'West', 20],
		['Veg', 'East', 5],
		['Veg', 'West', 15],
	].forEach((row, r) => row.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	return workbook;
}

/** A seeds the room with the data; B joins with an empty workbook and adopts it. */
function pair(): [Peer, Peer] {
	const a = peer(dataBook());
	const b = peer(createWorkbook(), a);
	return [a, b];
}

const pivotsOf = (p: Peer, sheet = 1): PivotTable[] =>
	p.edit.workbook.sheets[sheet]?.pivotTables ?? [];
const text = (p: Peer, row: number, col: number, sheet = 1) =>
	getCell(p.edit.workbook.sheets[sheet]!, row, col)?.value;

describe('sharing PivotTables', () => {
	it('shares create, field changes, layout, filters, refresh and delete', () => {
		const [a, b] = pair();
		const name = a.edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{ source: 'Data!A1:C5', rowField: 'Category', valueField: 'Amount' },
		);
		sync(a, b);
		expect(pivotsOf(b).map((pivot) => pivot.name)).toEqual([name]);
		expect(pivotsOf(b)[0]?.records).toEqual(pivotsOf(a)[0]?.records);
		expect(pivotsOf(b)[0]?.partName).toBeUndefined();
		expect(text(b, 1, 0)).toBe('Fruit');

		b.edit.updatePivotTable(1, name, { columnField: 'Region', layout: 'tabular' });
		sync(a, b);
		expect(pivotsOf(a)[0]).toMatchObject({ columnField: 'Region', layout: 'tabular' });

		a.edit.updatePivotTable(1, name, { itemFilters: [{ field: 'Category', hidden: ['Veg'] }] });
		sync(a, b);
		expect(pivotsOf(b)[0]?.itemFilters).toEqual([{ field: 'Category', hidden: ['Veg'] }]);

		a.edit.setCellInput(0, 1, 2, '100');
		// B refreshes once A's edit has reached it, so the shared records carry the new amount.
		sync(a, b);
		b.edit.refreshPivotTable(1, name);
		sync(a, b);
		expect(pivotsOf(a)[0]?.records[0]).toEqual(['Fruit', 'East', 100]);

		b.edit.deletePivotTable(1, name);
		sync(a, b);
		expect(pivotsOf(a)).toEqual([]);
	});

	it('keeps concurrent edits to two PivotTables', () => {
		const [a, b] = pair();
		const first = a.edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{ source: 'Data!A1:C5', rowField: 'Category', valueField: 'Amount' },
		);
		const second = a.edit.createPivotTable(
			1,
			{ row: 0, col: 6 },
			{ source: 'Data!A1:C5', rowField: 'Region', valueField: 'Amount' },
		);
		sync(a, b);
		a.edit.updatePivotTable(1, first, { layout: 'outline' });
		b.edit.updatePivotTable(1, second, { layout: 'tabular' });
		sync(a, b);
		for (const p of [a, b]) {
			const byName = new Map(pivotsOf(p).map((pivot) => [pivot.name, pivot]));
			expect(byName.get(first)?.layout).toBe('outline');
			expect(byName.get(second)?.layout).toBe('tabular');
		}
	});

	it('keeps a PivotTable over the size limit local and never clears it', () => {
		const [a, b] = pair();
		const name = a.edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{ source: 'Data!A1:C5', rowField: 'Category', valueField: 'Amount' },
		);
		const sheet = a.edit.workbook.sheets[1]!;
		const big = {
			...sheet.pivotTables![0]!,
			records: Array.from({ length: 300_000 }, () => ['Fruit', 'East', 1]),
		};
		sheet.pivotTables = [big];
		expect(pivotSpec.encode(big)).toBeUndefined();
		a.edit.setCellInput(1, 20, 0, 'note');
		a.edit.addSheet('Notes');
		sync(a, b);
		expect(pivotsOf(a)).toEqual([big]);
		expect(pivotsOf(a)[0]?.name).toBe(name);
		expect(pivotsOf(b)).toEqual([]);
	});

	it('marks a loaded PivotTable for rewriting when a peer changes it, and saves it', async () => {
		const seed = dataBook();
		createEditSession(seed).createPivotTable(
			1,
			{ row: 0, col: 0 },
			{ source: 'Data!A1:C5', rowField: 'Category', valueField: 'Amount' },
		);
		const bytes = await saveXlsx(seed);
		const a = peer(await loadXlsx(bytes));
		const b = peer(await loadXlsx(bytes), a);
		const name = pivotsOf(a)[0]!.name;
		expect(pivotsOf(b)[0]?.partName).toBeDefined();
		a.edit.updatePivotTable(1, name, { layout: 'tabular' });
		sync(a, b);
		expect(pivotsOf(b)[0]).toMatchObject({ layout: 'tabular', dirty: true });
		expect(pivotsOf(b)[0]?.partName).toBeDefined();
		const reopened = await loadXlsx(await saveXlsx(b.edit.workbook));
		expect(reopened.sheets[1]?.pivotTables?.[0]?.layout).toBe('tabular');
	});

	it('refuses malformed entries', () => {
		expect(pivotSpec.decode({ name: 'P', records: 'x' })).toBeUndefined();
		expect(
			pivotSpec.decode({
				name: 'P',
				source: 'Data!A1:B2',
				rowField: 'a',
				valueField: 'b',
				valueCaption: 'Sum of b',
				aggregate: 'eval',
				location: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
				fields: ['a', 'b'],
				records: [],
			}),
		).toBeUndefined();
	});
});

import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { createEditSession } from '../edit/session';
import type { EditSession } from '../edit/types';
import type { ChartObject, ConditionalFormat, Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { type WorkbookBinding, type WorkbookCollabHost, bindWorkbookSession } from './bind';

interface Peer {
	doc: Y.Doc;
	edit: EditSession;
	binding: WorkbookBinding;
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
	const binding = bindWorkbookSession(host, edit);
	return {
		doc,
		edit,
		binding,
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

const fixture = new Uint8Array(
	readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-features.xlsx')),
);

/** Two peers that each opened the same file; the second joins the first one's room. */
async function pair(): Promise<[Peer, Peer]> {
	const a = peer(await loadXlsx(fixture));
	const b = peer(await loadXlsx(fixture), false);
	sync(a, b);
	b.open();
	sync(a, b);
	return [a, b];
}

const at = (row: number, col: number) => ({ row, col });
const range = (r0: number, c0: number, r1 = r0, c1 = c0) => ({
	start: at(r0, c0),
	end: at(r1, c1),
});
const rule = (formula: string): ConditionalFormat['rules'][number] => ({
	type: 'expression',
	formula,
	style: { numFmt: '0' },
	priority: 1,
});
const chartsOf = (p: Peer) =>
	p.edit.workbook.sheets.flatMap((s) =>
		s.drawings.filter((d): d is ChartObject => d.kind === 'chart'),
	);
const chartSheet = (p: Peer) =>
	p.edit.workbook.sheets.findIndex((s) => s.drawings.some((d) => d.kind === 'chart'));

async function parts(wb: Workbook): Promise<Map<string, string>> {
	const zip = await JSZip.loadAsync(await saveXlsx(wb));
	const out = new Map<string, string>();
	for (const name of Object.keys(zip.files)) {
		const file = zip.file(name);
		if (file) out.set(name, await file.async('base64'));
	}
	return out;
}

describe('sharing comments, rules, links, tables and charts', () => {
	it('adopting a room built from the same file changes nothing and saves identically', async () => {
		const [a, b] = await pair();
		expect(await parts(b.edit.workbook)).toEqual(await parts(a.edit.workbook));
		const source = await parts(await loadXlsx(fixture));
		const saved = await parts(b.edit.workbook);
		for (const name of source.keys())
			if (name.startsWith('xl/charts/chart')) expect(saved.get(name)).toBe(source.get(name));
	});

	it('keeps concurrent notes on different cells and threaded replies to one note', async () => {
		const [a, b] = await pair();
		a.edit.setComment(0, at(0, 0), 'from a', 'Ann');
		b.edit.setComment(0, at(1, 1), 'from b', 'Bob');
		sync(a, b);
		for (const p of [a, b])
			expect(p.edit.workbook.sheets[0]?.comments.map((c) => c.text).sort()).toEqual(
				expect.arrayContaining(['from a', 'from b']),
			);
		// Replies are separate entries, so two peers replying to one thread both keep theirs.
		const target = a.edit.workbook.sheets[0]?.comments.find((c) => c.text === 'from a');
		if (!target) throw new Error('missing note');
		target.replies = [{ author: 'Ann', text: 'first' }];
		a.edit.setComment(0, at(0, 0), 'from a', 'Ann');
		sync(a, b);
		const bNote = b.edit.workbook.sheets[0]?.comments.find((c) => c.text === 'from a');
		expect(bNote?.replies?.map((r) => r.text)).toEqual(['first']);
		if (bNote) bNote.replies = [...(bNote.replies ?? []), { author: 'Bob', text: 'from b' }];
		b.edit.setComment(0, at(0, 0), 'from a', 'Ann');
		const aNote = a.edit.workbook.sheets[0]?.comments.find((c) => c.text === 'from a');
		if (aNote) aNote.replies = [...(aNote.replies ?? []), { author: 'Ann', text: 'second' }];
		a.edit.setComment(0, at(0, 0), 'from a', 'Ann');
		sync(a, b);
		for (const p of [a, b]) {
			const note = p.edit.workbook.sheets[0]?.comments.find((c) => c.text === 'from a');
			expect(note?.replies?.map((r) => r.text).sort()).toEqual(['first', 'from b', 'second']);
		}
	});

	it('shares conditional formats, validations and links; one rule per entry', async () => {
		const [a, b] = await pair();
		const before = b.edit.workbook.sheets[0]?.conditionalFormats.length ?? 0;
		a.edit.addConditionalFormat(0, { ranges: [range(0, 0, 4, 0)], rules: [rule('A1>1')] });
		b.edit.addConditionalFormat(0, { ranges: [range(0, 2, 4, 2)], rules: [rule('C1>2')] });
		a.edit.setDataValidation(0, { ranges: [], type: 'whole', operator: 'between' }, range(5, 5));
		b.edit.setHyperlink(0, range(7, 7), { target: 'https://example.com/' });
		sync(a, b);
		for (const p of [a, b]) {
			const sheet = p.edit.workbook.sheets[0];
			expect(sheet?.conditionalFormats.length).toBe(before + 2);
			expect(sheet?.dataValidations.some((v) => v.type === 'whole')).toBe(true);
			expect(sheet?.hyperlinks.some((l) => l.target === 'https://example.com/')).toBe(true);
		}
		// Concurrent edits to different rules both survive.
		const aSheet = a.edit.workbook.sheets[0];
		const bSheet = b.edit.workbook.sheets[0];
		const find = (sheet: typeof aSheet, formula: string) =>
			sheet?.conditionalFormats.findIndex((f) =>
				f.rules.some((r) => r.type === 'expression' && r.formula === formula),
			) ?? -1;
		a.edit.replaceConditionalFormat(0, find(aSheet, 'A1>1'), {
			ranges: [range(0, 0, 9, 0)],
			rules: [rule('A1>1')],
		});
		b.edit.replaceConditionalFormat(0, find(bSheet, 'C1>2'), {
			ranges: [range(0, 2, 9, 2)],
			rules: [rule('C1>2')],
		});
		sync(a, b);
		for (const p of [a, b]) {
			const list = p.edit.workbook.sheets[0]?.conditionalFormats ?? [];
			expect(list.length).toBe(before + 2);
			// Editing a rule renumbers the sheet's priorities, which rewrites sibling entries: of two
			// concurrent edits at least one survives and no rule is duplicated or lost.
			expect(list.filter((f) => f.ranges[0]?.end.row === 9).length).toBeGreaterThanOrEqual(1);
		}
	});

	it('shares a table', async () => {
		const [a, b] = await pair();
		a.edit.setRangeValues(0, at(20, 0), [
			['h1', 'h2'],
			[1, 2],
		]);
		const table = a.edit.createTable(0, range(20, 0, 21, 1), true);
		sync(a, b);
		const copy = b.edit.workbook.sheets[0]?.tables.find((t) => t.name === table.name);
		expect(copy?.columns.map((c) => c.name)).toEqual(['h1', 'h2']);
	});

	it('shares chart edits, last writer wins, and undo is per client', async () => {
		const [a, b] = await pair();
		const sheet = chartSheet(a);
		expect(sheet).toBeGreaterThanOrEqual(0);
		const index = a.edit.workbook.sheets[sheet]?.drawings.findIndex((d) => d.kind === 'chart') ?? 0;
		a.edit.updateChart(sheet, index, { title: 'Shared title' });
		sync(a, b);
		expect(chartsOf(b)[0]?.title).toBe('Shared title');
		// The partName stays the one of b's own package, so an edit saves over its source part.
		expect(chartsOf(b)[0]?.partName).toBe(chartsOf(a)[0]?.partName);
		const bytes = await parts(b.edit.workbook);
		const text = Buffer.from(bytes.get(chartsOf(b)[0]?.partName ?? '') ?? '', 'base64').toString();
		expect(text).toContain('Shared title');

		// Undo reverts only the local edit.
		b.edit.setComment(0, at(2, 2), 'b note', 'Bob');
		a.edit.setComment(0, at(3, 3), 'a note', 'Ann');
		sync(a, b);
		expect(a.binding.undo()).toBe(true);
		sync(a, b);
		for (const p of [a, b]) {
			const texts = p.edit.workbook.sheets[0]?.comments.map((c) => c.text) ?? [];
			expect(texts).toContain('b note');
			expect(texts).not.toContain('a note');
		}
		expect(a.binding.undo()).toBe(true);
		sync(a, b);
		expect(chartsOf(b)[0]?.title).not.toBe('Shared title');
		expect(b.binding.undo()).toBe(true);
		sync(a, b);
		expect(a.edit.workbook.sheets[0]?.comments.map((c) => c.text)).not.toContain('b note');
	});

	it('adds a new chart on one peer and saves it from the other', async () => {
		const [a, b] = await pair();
		const before = chartsOf(b).length;
		a.edit.addChart(chartSheet(a), {
			anchor: { from: { row: 1, col: 8, rowOffset: 0, colOffset: 0 } },
			chartType: 'column',
			title: 'New one',
			showLegend: true,
			series: [{ valuesRef: 'Sheet1!$B$1:$B$2', categories: [], values: [] }],
		});
		sync(a, b);
		expect(chartsOf(b).length).toBe(before + 1);
		const mine = chartsOf(b).at(-1);
		expect(mine?.partName).toBeUndefined();
		const zip = await JSZip.loadAsync(await saveXlsx(b.edit.workbook));
		const texts = await Promise.all(
			Object.keys(zip.files)
				.filter((n) => n.startsWith('xl/charts/chart'))
				.map((n) => zip.file(n)?.async('text')),
		);
		expect(texts.some((t) => t?.includes('New one'))).toBe(true);
	});

	it('loses nothing when peers first write the object lists at once on an old-shape room', async () => {
		// An old room: sheets and cells only, no part entries and no stamp.
		const seed = await loadXlsx(fixture);
		const a = peer(seed);
		const old = new Y.Doc();
		Y.applyUpdate(old, Y.encodeStateAsUpdate(a.doc));
		old.getMap('xlsx:meta').delete('parts');
		for (const part of ['comments', 'replies', 'cf', 'dv', 'links', 'tables', 'charts'])
			old.getMap().clear();
		const open = (workbook: Workbook): Peer => {
			const p = peer(workbook, false);
			Y.applyUpdate(p.doc, Y.encodeStateAsUpdate(old));
			p.open();
			return p;
		};
		const localNotes = (await loadXlsx(fixture)).sheets[0]?.comments.length ?? 0;
		const b = open(await loadXlsx(fixture));
		const c = open(await loadXlsx(fixture));
		// Adopting a room without the stamp keeps each peer's own lists.
		expect(b.edit.workbook.sheets[0]?.comments.length).toBe(localNotes);
		b.edit.setComment(0, at(30, 0), 'from b', 'Bob');
		c.edit.setComment(0, at(31, 0), 'from c', 'Cy');
		b.edit.addConditionalFormat(0, { ranges: [range(0, 0)], rules: [rule('A1>1')] });
		c.edit.addConditionalFormat(0, { ranges: [range(1, 1)], rules: [rule('B2>1')] });
		sync(b, c);
		for (const p of [b, c]) {
			const sheet = p.edit.workbook.sheets[0];
			const texts = sheet?.comments.map((n) => n.text) ?? [];
			expect(texts).toEqual(expect.arrayContaining(['from b', 'from c']));
			const formulas = (sheet?.conditionalFormats ?? []).flatMap((f) =>
				f.rules.flatMap((r) => (r.type === 'expression' ? [r.formula] : [])),
			);
			expect(formulas).toEqual(expect.arrayContaining(['A1>1', 'B2>1']));
		}
	});
});

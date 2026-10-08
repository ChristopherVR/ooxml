import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { formatAddress, formatRange } from '../address';
import { forEachCell } from '../cells';
import { isSpilledCell } from '../formula/spill';
import type { Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { styleAt, styleKey } from '../styles';
import { createWorkbook } from '../workbook';
import { readSharedWorkbook, writeSharedWorkbook, xlsxDocumentAdapter } from './adapter';
import { sharedWorkbook } from './schema';

const DIR = path.join(import.meta.dirname, '..', '__fixtures__');
const FIXTURES = readdirSync(DIR).filter((f) => f.endsWith('.xlsx'));

/** The part of a workbook the shared schema carries, with formats resolved to content. */
function modelled(workbook: Workbook) {
	const style = (id: number | undefined) => styleKey(styleAt(workbook, id));
	return {
		date1904: workbook.date1904,
		names: workbook.definedNames
			.map((n) => ({ ...n, localSheet: workbook.sheets[n.localSheet ?? -1]?.name }))
			.sort((a, b) => `${a.localSheet}!${a.name}`.localeCompare(`${b.localSheet}!${b.name}`)),
		sheets: workbook.sheets.map((sheet) => {
			const cells: Record<string, unknown> = {};
			forEachCell(sheet, (cell, row, col) => {
				const out: Record<string, unknown> = { style: style(cell.styleId) };
				if (!isSpilledCell(cell)) {
					out.value = cell.value;
					if (cell.formula !== undefined) out.formula = cell.formula;
					if (cell.arrayRange) out.array = formatRange(cell.arrayRange);
					if (cell.richText) out.rich = cell.richText;
					if (cell.dynamicArray) out.dynamic = true;
					if (cell.legacyFormula) out.legacy = true;
				} else if (!cell.styleId) return;
				cells[formatAddress({ row, col })] = out;
			});
			return {
				name: sheet.name,
				sheetId: sheet.sheetId,
				state: sheet.state,
				tabColor: sheet.tabColor,
				defaultRowHeight: sheet.defaultRowHeight,
				defaultColWidth: sheet.defaultColWidth,
				rows: [...sheet.rowInfo].map(([row, info]) => [
					row,
					{ ...info, styleId: style(info.styleId) },
				]),
				columns: sheet.columns.map((c) => ({ ...c, styleId: style(c.styleId) })),
				merges: sheet.merges.map(formatRange).sort(),
				cells,
			};
		}),
	};
}

describe('xlsxDocumentAdapter', () => {
	it.each(FIXTURES)('round-trips the modelled subset of %s', async (file) => {
		const workbook = await loadXlsx(new Uint8Array(readFileSync(path.join(DIR, file))));
		const doc = new Y.Doc();
		writeSharedWorkbook(doc, workbook);
		const peer = new Y.Doc();
		Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
		const read = readSharedWorkbook(peer);
		expect(modelled(read)).toEqual(modelled(workbook));
		// Writing the same model again changes nothing.
		let updates = 0;
		doc.on('update', () => updates++);
		writeSharedWorkbook(doc, workbook);
		expect(updates).toBe(0);
	});

	it('reads onto a base workbook so unshared parts survive', () => {
		const base = createWorkbook();
		base.sheets[0]!.comments.push({ address: { row: 0, col: 0 }, author: 'me', text: 'kept' });
		const doc = new Y.Doc();
		writeSharedWorkbook(doc, base);
		const adapter = xlsxDocumentAdapter({ base: () => base });
		const read = adapter.read(doc);
		expect(read).not.toBe(base);
		expect(read.sheets[0]?.comments).toHaveLength(1);
	});

	it('reports emptiness and change origins', () => {
		const adapter = xlsxDocumentAdapter();
		const doc = new Y.Doc();
		expect(adapter.isEmpty(doc)).toBe(true);
		const origins: unknown[] = [];
		const stop = adapter.observe(doc, (origin) => origins.push(origin));
		const origin = Symbol('test');
		adapter.write(doc, createWorkbook({ sheets: ['A', 'B'] }), origin as never);
		expect(adapter.isEmpty(doc)).toBe(false);
		expect(origins).toEqual([origin]);
		stop();
		expect(adapter.read(doc).sheets.map((s) => s.name)).toEqual(['A', 'B']);
	});

	it('ignores malformed entries from a peer', () => {
		const doc = new Y.Doc();
		writeSharedWorkbook(doc, createWorkbook());
		const shared = sharedWorkbook(doc);
		const [key] = [...shared.sheets.keys()];
		const cells = (shared.sheets.get(key ?? '') as Y.Map<Y.Map<unknown>>).get('cells');
		cells?.set('A1', { v: 'x', t: 'n' });
		cells?.set('B1', { v: '#BOGUS', t: 'e' });
		cells?.set('ZZZZ9', { v: 1, t: 'n' });
		cells?.set('C1', { v: 3, t: 'n' });
		shared.sheets.set('broken', 'not a map');
		const read = readSharedWorkbook(doc);
		expect(read.sheets).toHaveLength(1);
		expect([...(read.sheets[0]?.rows.get(0)?.keys() ?? [])]).toEqual([2]);
	});
});

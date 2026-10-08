import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { createWorkbook } from '../workbook';
import { applyColumns } from './lines';
import { resolveSelections, sanitizeXlsxPresence } from './presence';
import { reorder, sheetEntries } from './sheets';
import { newSheetMap } from './schema';

describe('reorder', () => {
	it('moves as few sheets as possible', () => {
		expect([...reorder([3, 1, 2])]).toEqual([[0, 0.5]]);
		expect([...reorder([1, 2, undefined])]).toEqual([[2, 3]]);
		expect([...reorder([1, undefined, 2])]).toEqual([[1, 1.5]]);
		expect(reorder([1, 2, 3]).size).toBe(0);
	});

	it('renumbers when neighbours run out of precision', () => {
		expect([...reorder([1, undefined, 1 + Number.EPSILON])]).toEqual([
			[0, 1],
			[1, 2],
			[2, 3],
		]);
	});
});

describe('sheetEntries', () => {
	it('orders sheets and makes colliding names and ids unique by key', () => {
		const doc = new Y.Doc();
		const sheets = doc.getMap<unknown>('s');
		const add = (key: string, name: string, sheetId: number, order: number) => {
			const map = newSheetMap();
			sheets.set(key, map);
			const props = map.get('props') as Y.Map<unknown>;
			props.set('name', name);
			props.set('sheetId', sheetId);
			props.set('order', order);
		};
		add('b-0', 'Data', 2, 1);
		add('a-0', 'data', 2, 2);
		add('c-0', 'Other', 1, 0.5);
		const entries = sheetEntries(sheets);
		expect(entries.map((e) => [e.key, e.name, e.sheetId])).toEqual([
			['c-0', 'Other', 1],
			['b-0', 'Data (2)', 3],
			['a-0', 'data', 2],
		]);
	});
});

describe('applyColumns', () => {
	it('keeps non-overlapping spans intact and resolves overlaps to the narrowest', () => {
		const doc = new Y.Doc();
		const cols = doc.getMap<unknown>('c');
		cols.set('A:B', { size: 10 });
		cols.set('C:C', { size: 20 });
		cols.set('A:Z', { size: 5 });
		const sheet = createWorkbook().sheets[0]!;
		applyColumns(sheet, cols, () => undefined);
		expect(sheet.columns.map((c) => [c.min, c.max, c.width])).toEqual([
			[0, 1, 10],
			[2, 2, 20],
			[3, 25, 5],
		]);
	});
});

describe('presence', () => {
	it('sanitises payloads and resolves sheet keys', () => {
		expect(sanitizeXlsxPresence({ sheet: 'k1-0', range: 'b2:c4' })).toEqual({
			sheet: 'k1-0',
			range: 'B2:C4',
		});
		expect(sanitizeXlsxPresence({ sheet: '<img>', range: 'A1' })).toEqual({});
		expect(sanitizeXlsxPresence({ sheet: 'k1-0', range: 'nope' })).toEqual({});
		const peers = [
			{
				clientId: 7,
				userName: 'Ann',
				userColor: '#f00',
				lastUpdated: '',
				sheet: 'k1-0',
				range: 'A1',
			},
			{
				clientId: 8,
				userName: 'Bob',
				userColor: '#0f0',
				lastUpdated: '',
				sheet: 'gone-1',
				range: 'A1',
			},
		];
		expect(resolveSelections(peers, (key) => (key === 'k1-0' ? 2 : undefined))).toEqual([
			{
				clientId: 7,
				userName: 'Ann',
				userColor: '#f00',
				sheet: 2,
				range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
			},
		]);
	});
});

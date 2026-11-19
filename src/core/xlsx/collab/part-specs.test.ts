import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';

import { LEFT_OUT_MESSAGE, scanUpdate, MIN_BYTES_PER_STRUCT } from '../../collab/index';
import { putCell } from '../cells';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { readSharedWorkbook, writeSharedWorkbook } from './adapter';
import { linkSpec } from './part-specs';
import { sharedWorkbook } from './schema';

describe('linkSpec.decode', () => {
	it('keeps the link fields a document holds', () => {
		expect(
			linkSpec.decode({
				r: 'A1',
				target: 'https://example.com/',
				location: 'Sheet2!A1',
				tooltip: 'Open',
				display: 'Example',
			}),
		).toEqual({
			range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
			target: 'https://example.com/',
			location: 'Sheet2!A1',
			tooltip: 'Open',
			display: 'Example',
		});
	});

	it('drops a script target and any field that is not text, and names them', () => {
		const refused: string[] = [];
		const link = linkSpec.decode(
			{ r: 'B2', target: 'javascript:alert(1)', tooltip: { toString: 'x' }, extra: '<b>' },
			(field) => refused.push(field),
		);
		expect(link).toEqual({ range: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } } });
		expect(refused).toEqual(['target', 'tooltip']);
	});
});

// Every kind of target a workbook can hold, as the loader keeps them.
const TARGETS = [
	'https://example.com/a?b=1#c',
	'mailto:a@example.com',
	'tel:+15551234',
	'sip:alice@example.com',
	'msteams:/l/chat/0/0?users=a@example.com',
	'onenote:https://example.com/notes.one',
	'ftp://files.example.com/a.txt',
	'file:///C:/Reports/q1.xlsx',
	'Book2.xlsx',
	'../Docs/report.docx',
	'\\\\server\\share\\y.xlsx',
];

describe('hyperlinks shared with a peer', () => {
	it('reach the peer with every target the loader keeps', async () => {
		const source = createWorkbook();
		TARGETS.forEach((target, row) => {
			putCell(source.sheets[0]!, row, 0, { value: target });
			source.sheets[0]!.hyperlinks.push({
				range: { start: { row, col: 0 }, end: { row, col: 0 } },
				target,
			});
		});
		// The targets survive a save and load, so these are what a loaded workbook holds.
		const loaded = await loadXlsx(await saveXlsx(source));
		const targets = loaded.sheets[0]!.hyperlinks.map((link) => link.target);
		expect(targets).toEqual(TARGETS);

		const a = new Y.Doc();
		writeSharedWorkbook(a, loaded);
		const b = new Y.Doc();
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		const peer = readSharedWorkbook(b);
		expect(peer.sheets[0]!.hyperlinks.map((link) => link.target)).toEqual(TARGETS);
		expect(peer.warnings).not.toContain(LEFT_OUT_MESSAGE);
	});

	it('leave out only a script target, and say so', () => {
		const a = new Y.Doc();
		const source = createWorkbook();
		source.sheets[0]!.hyperlinks.push({
			range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
			target: 'tel:+15551234',
		});
		writeSharedWorkbook(a, source);
		const shared = sharedWorkbook(a);
		const [sheetKey] = [...shared.sheets.keys()];
		shared.parts.links.set(`${sheetKey}/B1`, { i: 1, d: { r: 'B1', target: 'javascript:x()' } });
		const peer = readSharedWorkbook(a);
		expect(peer.sheets[0]!.hyperlinks).toEqual([
			{ range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }, target: 'tel:+15551234' },
			{ range: { start: { row: 0, col: 1 }, end: { row: 0, col: 1 } } },
		]);
		expect(peer.warnings).toContain(LEFT_OUT_MESSAGE);
	});
});

describe('the update budget against real workbooks', () => {
	it('admits the first sync of the densest sheet at less than half its budget', () => {
		// Boolean cells are the densest content measured (see docs/collab-area.md).
		const workbook = createWorkbook();
		for (let row = 0; row < 20_000; row++)
			for (let col = 0; col < 10; col++) putCell(workbook.sheets[0]!, row, col, { value: true });
		const doc = new Y.Doc();
		writeSharedWorkbook(doc, workbook);
		const update = Y.encodeStateAsUpdate(doc);
		const scan = scanUpdate(update);
		expect(scan.ok).toBe(true);
		if (scan.ok) expect(scan.structs * MIN_BYTES_PER_STRUCT * 2).toBeLessThan(update.byteLength);
	});
});

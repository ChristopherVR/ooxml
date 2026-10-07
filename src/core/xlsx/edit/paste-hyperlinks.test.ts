import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseRange, rangeContains } from '../address.js';
import { getCell } from '../cells.js';
import type { Hyperlink } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from '../write/index.js';
import { createEditSession } from './session.js';
import type { PasteOptions } from './types.js';

interface NativeLink {
	target: string;
	location: string;
	tooltip: string;
}
interface NativeCase extends Required<PasteOptions> {
	cells: { row: number; col: number; value: number | null; link: NativeLink | null }[];
}
const fixture = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-paste-hyperlinks.json', import.meta.url), 'utf8'),
) as { cases: NativeCase[] };
const range = (ref: string) => parseRange(ref)!;
const setup = () => {
	const workbook = createWorkbook();
	const sheet = workbook.sheets[0]!;
	const session = createEditSession(workbook);
	session.setCellValue(0, 0, 0, 2);
	session.setCellValue(0, 0, 1, 4);
	session.setCellInput(0, 1, 0, '=B2+$J$1');
	session.setCellValue(0, 0, 9, 3);
	sheet.hyperlinks = [
		{ range: range('A1'), target: 'https://example.com/source?x=1&y=2', tooltip: 'source tip' },
		{ range: range('B1'), location: 'Sheet1!$J$1', tooltip: 'internal tip' },
	];
	for (const ref of ['D4', 'E4', 'D5', 'E5']) {
		const at = range(ref).start;
		session.setCellValue(0, at.row, at.col, 9);
		sheet.hyperlinks.push({
			range: range(ref),
			target: 'https://example.com/destination',
			tooltip: 'dest tip',
		});
	}
	return { workbook, sheet, session };
};
const normalized = (link: Hyperlink | undefined): NativeLink | null =>
	link
		? { target: link.target ?? '', location: link.location ?? '', tooltip: link.tooltip ?? '' }
		: null;

describe('Hyperlink clipboard semantics recorded in Microsoft Excel', () => {
	it.each(fixture.cases)('$mode transpose=$transpose skip=$skipBlanks op=$operation', (native) => {
		const { sheet, session } = setup();
		const payload = session.copy(0, range('A1:B2'));
		const before = structuredClone(sheet.hyperlinks);
		session.paste(0, range('D4:E5'), payload, native);
		for (const cell of native.cells) {
			expect(getCell(sheet, cell.row, cell.col)?.value ?? null).toBe(cell.value);
			expect(normalized(sheet.hyperlinks.find((h) => rangeContains(h.range, cell)))).toEqual(
				cell.link,
			);
		}
		const after = structuredClone(sheet.hyperlinks);
		session.undo();
		expect(sheet.hyperlinks).toEqual(before);
		session.redo();
		expect(sheet.hyperlinks).toEqual(after);
	});
	it('clips hyperlink spans, splits destination spans, tiles snapshots and saves their targets', async () => {
		const { sheet, session } = setup();
		sheet.hyperlinks = [
			{ range: range('A1:C3'), target: 'https://example.com/source', tooltip: 'tip' },
		];
		const payload = session.copy(0, range('B2:C3'));
		sheet.hyperlinks = [];
		const target = createWorkbook();
		const dest = target.sheets[0]!;
		dest.hyperlinks = [{ range: range('D4:I9'), target: 'https://example.com/dest' }];
		const edits = createEditSession(target);
		edits.paste(0, range('D4:G7'), payload, { mode: 'all', transpose: true });
		const loaded = await loadXlsx(await saveXlsx(target));
		for (const [ref, url] of [
			['D4', 'source'],
			['G7', 'source'],
			['H4', 'dest'],
			['D8', 'dest'],
		])
			expect(
				loaded.sheets[0]!.hyperlinks.find((h) => rangeContains(h.range, range(ref!).start))?.target,
			).toBe(`https://example.com/${url}`);
		edits.undo();
		expect(dest.hyperlinks).toEqual([
			{ range: range('D4:I9'), target: 'https://example.com/dest' },
		]);
	});
	it('moves hyperlink anchors and their in-workbook targets with cut and restores them on undo', () => {
		const { workbook, sheet, session } = setup();
		sheet.hyperlinks = [{ range: range('A1:B2'), location: 'Sheet1!A1', tooltip: 'move target' }];
		const before = structuredClone(sheet.hyperlinks);
		const target = session.addSheet('Target');
		session.paste(target, range('D4'), session.cut(0, range('A1:B2')));
		expect(sheet.hyperlinks).toEqual([]);
		expect(workbook.sheets[target]!.hyperlinks).toEqual([
			{ range: range('D4:E5'), location: 'Target!D4', tooltip: 'move target' },
		]);
		session.undo();
		expect(sheet.hyperlinks).toEqual(before);
		expect(workbook.sheets[target]!.hyperlinks).toEqual([]);
	});
	it('round-trips external and in-workbook anchors through HTML clipboard text', () => {
		const { workbook, session } = setup();
		const payload = session.copy(0, range('A1:B1'));
		expect(payload.html).toContain('href="https://example.com/source?x=1&amp;y=2"');
		expect(payload.html).toContain('href="#Sheet1!$J$1"');
		session.paste(0, range('G7'), payload.html);
		expect(
			normalized(
				workbook.sheets[0]!.hyperlinks.find((h) => rangeContains(h.range, range('G7').start)),
			),
		).toEqual({
			target: 'https://example.com/source?x=1&y=2',
			location: '',
			tooltip: 'source tip',
		});
		expect(
			normalized(
				workbook.sheets[0]!.hyperlinks.find((h) => rangeContains(h.range, range('H7').start)),
			),
		).toEqual({ target: '', location: 'Sheet1!$J$1', tooltip: 'internal tip' });
	});
	it('reuses shared hyperlink policy for untrusted HTML without introducing active markup', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		session.paste(
			0,
			range('A1'),
			'<table><tr><td><a href="java&#x73;cript:alert(1)">label</a></td><td><a href="mailto:a@example.com" title="&quot;tip&quot;">mail</a></td></tr></table>',
		);
		expect(workbook.sheets[0]!.hyperlinks).toEqual([
			{ range: range('B1'), target: 'mailto:a@example.com', tooltip: '"tip"' },
		]);
		expect(getCell(workbook.sheets[0]!, 0, 0)?.value).toBe('label');
	});
});

/**
 * A table cell made in code takes its own edge lines, with a preset dash or
 * an exact px pattern written as `a:custDash`. Each saved case is loaded
 * back, so the written `a:tcPr` and the reloaded cell style agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxTableCellStyle, TablePptxElement } from '../../types';
import { createTableElement } from './ElementFactory';
import type { TableCellBordersInput, TableCellInput, TableInput } from './types';

const oneCell = (cell: Partial<TableCellInput>): TableInput => ({
	style: 'none',
	rows: [{ cells: [{ text: 'cell', ...cell }] }],
});

async function slideXml(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
}

async function roundTrip(input: TableInput): Promise<{
	tcPr: string;
	style: PptxTableCellStyle | undefined;
	resaved: string;
}> {
	const { handler, createSlide } = await PptxHandler.create();
	const bytes = await handler.save([createSlide('Blank').addTable(input).build()]);
	const xml = await slideXml(bytes);
	const reloader = new PptxHandler();
	const loaded = await reloader.load(bytes);
	const table = loaded.slides[0]?.elements.find((e) => e.type === 'table') as
		| TablePptxElement
		| undefined;
	const resaved = await slideXml(await reloader.save(loaded.slides));
	return {
		tcPr: /<a:tcPr\b[\s\S]*?<\/a:tcPr>/u.exec(xml)?.[0] ?? '',
		style: table?.tableData?.rows[0]?.cells[0]?.style,
		resaved: /<a:tcPr\b[\s\S]*?<\/a:tcPr>/u.exec(resaved)?.[0] ?? '',
	};
}

/** The `a:ds` pairs of one edge line. */
const segmentsOf = (tcPr: string, edge: string): string[] => {
	const ln = new RegExp(`<a:${edge}\\b[\\s\\S]*?</a:${edge}>`, 'u').exec(tcPr)?.[0] ?? '';
	return [...ln.matchAll(/<a:ds d="(\d+)" sp="(\d+)"/gu)].map((m) => `${m[1]}/${m[2]}`);
};

const children = (tcPr: string): string[] =>
	[...tcPr.matchAll(/<(a:\w+)\b/gu)].map((m) => m[1] ?? '').filter((n) => n !== 'a:tcPr');

describe('TableCellInput.borders', () => {
	it('writes a preset dash on every edge and reads it back', async () => {
		const { tcPr, style } = await roundTrip(
			oneCell({ borders: { all: { color: '#7F7F7F', width: 1, dash: 'sysDot' } } }),
		);
		for (const edge of ['lnL', 'lnR', 'lnT', 'lnB']) {
			expect(tcPr).toMatch(new RegExp(`<a:${edge} w="9525">[\\s\\S]*?<a:prstDash val="sysDot"`));
		}
		expect(tcPr).not.toContain('custDash');
		expect(style).toMatchObject({
			borderTopDash: 'sysDot',
			borderRightDash: 'sysDot',
			borderBottomDash: 'sysDot',
			borderLeftDash: 'sysDot',
			borderTopColor: '#7F7F7F',
		});
	});

	it('writes a px pattern as a:custDash relative to the width, and keeps it on resave', async () => {
		const { tcPr, style, resaved } = await roundTrip(
			oneCell({ borders: { all: { color: '#FF0000', width: 2, dash: [6, 2, 1, 2] } } }),
		);
		for (const edge of ['lnL', 'lnR', 'lnT', 'lnB']) {
			expect(segmentsOf(tcPr, edge)).toEqual(['300000/100000', '50000/100000']);
		}
		expect(tcPr).not.toContain('prstDash');
		expect(style?.borderLeftDash).toBe('custom');
		expect(style?.borderLeftDashSegments).toEqual([
			{ dash: 300000, space: 100000 },
			{ dash: 50000, space: 100000 },
		]);
		expect(segmentsOf(resaved, 'lnB')).toEqual(['300000/100000', '50000/100000']);
	});

	it('repeats an odd-length pattern once, as stroke-dasharray does', async () => {
		const { tcPr } = await roundTrip(
			oneCell({ borders: { top: { width: 1, dash: [4] }, left: { width: 1, dash: [3, 1, 1] } } }),
		);
		expect(segmentsOf(tcPr, 'lnT')).toEqual(['400000/400000']);
		expect(segmentsOf(tcPr, 'lnL')).toEqual(['300000/100000', '100000/300000', '100000/100000']);
	});

	it('measures a pattern against 1 px when the width is omitted or 0', async () => {
		const { tcPr, style } = await roundTrip(
			oneCell({ borders: { top: { dash: [2, 1] }, bottom: { width: 0, dash: [2, 1] } } }),
		);
		expect(segmentsOf(tcPr, 'lnT')).toEqual(['200000/100000']);
		expect(segmentsOf(tcPr, 'lnB')).toEqual(['200000/100000']);
		expect(tcPr).toMatch(/<a:lnB w="0"/u);
		expect(style?.borderTopWidth).toBeUndefined();
	});

	it('lets an edge override all, and writes only the edges given', async () => {
		const { tcPr, style } = await roundTrip(
			oneCell({
				borders: { all: { color: '#000000', width: 1, dash: 'dash' }, top: { dash: [8, 4] } },
			}),
		);
		expect(segmentsOf(tcPr, 'lnT')).toEqual(['800000/400000']);
		expect(tcPr).toMatch(/<a:lnT w="9525"><a:solidFill><a:srgbClr val="000000"/u);
		expect(style).toMatchObject({ borderTopDash: 'custom', borderBottomDash: 'dash' });

		const only = await roundTrip(oneCell({ borders: { bottom: { color: '#00B050' } } }));
		expect(children(only.tcPr).filter((n) => n.startsWith('a:ln'))).toEqual(['a:lnB']);
	});

	it('writes the edges, the dash and the fill in schema order', async () => {
		const { tcPr } = await roundTrip(
			oneCell({
				fill: { type: 'solid', color: '#DDEBF7' },
				borders: { all: { color: '#1F4E79', width: 1, dash: [3, 1] } },
			}),
		);
		const order = children(tcPr).filter((n) =>
			['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB', 'a:solidFill', 'a:custDash'].includes(n),
		);
		// CT_TableCellProperties: lnL, lnR, lnT, lnB, then the fill; in each
		// CT_LineProperties the line's fill comes before its dash.
		expect(order).toEqual([
			...['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'].flatMap((ln) => [ln, 'a:solidFill', 'a:custDash']),
			'a:solidFill',
		]);
	});

	it('refuses a pattern or width it cannot write', () => {
		const refused: TableCellBordersInput[] = [
			{ all: { dash: [] } },
			{ all: { dash: [0, 2] } },
			{ all: { dash: [2, 0] } },
			{ top: { dash: [-1, 2] } },
			{ top: { dash: [Number.NaN, 2] } },
			{ top: { dash: [Number.POSITIVE_INFINITY, 2] } },
			{ top: { width: 1, dash: [1e-9, 1] } },
			{ top: { width: 1, dash: [1e5, 1] } },
			{ left: { width: -1 } },
			{ left: { width: Number.NaN } },
			{ left: { dash: 'custom' as never } },
		];
		for (const borders of refused) {
			expect(() => createTableElement(oneCell({ borders })), JSON.stringify(borders)).toThrow(
				RangeError,
			);
		}
	});
});

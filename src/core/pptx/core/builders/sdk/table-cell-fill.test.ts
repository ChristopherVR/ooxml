/**
 * A table cell made in code keeps its own `fill`: it is written into the
 * cell's `a:tcPr` and read back as the same cell style. Fills a cell cannot
 * carry are refused. Each saved case is loaded back, so the written `a:tc`
 * and the reloaded cell agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxTableCell, TablePptxElement } from '../../types';
import { createTableElement } from './ElementFactory';
import type { FillInput, TableInput } from './types';

async function roundTrip(input: TableInput): Promise<{ cells: string[]; loaded: PptxTableCell[] }> {
	const { handler, createSlide } = await PptxHandler.create();
	const bytes = await handler.save([createSlide('Blank').addTable(input).build()]);
	const zip = await JSZip.loadAsync(bytes);
	const slide = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const reloaded = await new PptxHandler().load(bytes);
	const table = reloaded.slides[0]?.elements.find((e) => e.type === 'table') as
		| TablePptxElement
		| undefined;
	return {
		cells: slide.match(/<a:tc\b[\s\S]*?<\/a:tc>/gu) ?? [],
		loaded: table?.tableData?.rows.flatMap((row) => row.cells) ?? [],
	};
}

/** One row: a cell with the fill, then a plain cell. */
const withFill = (fill: FillInput, style?: string): TableInput => ({
	...(style !== undefined ? { style } : {}),
	rows: [{ cells: [{ text: 'filled', fill }, { text: 'plain' }] }],
});

const tcPr = (cell: string | undefined): string =>
	/<a:tcPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:tcPr>)/u.exec(cell ?? '')?.[0] ?? '';

describe('TableCellInput.fill', () => {
	it('writes a solid fill and reads it back; a cell without fill gets none', async () => {
		const { cells, loaded } = await roundTrip(withFill({ type: 'solid', color: '#FF0000' }));
		expect(tcPr(cells[0])).toMatch(/<a:solidFill><a:srgbClr val="FF0000"/u);
		expect(tcPr(cells[1])).not.toMatch(/Fill/u);
		expect(loaded[0]?.style).toMatchObject({ fillMode: 'solid', backgroundColor: '#FF0000' });
		expect(loaded[1]?.style?.fillMode).toBeUndefined();
	});

	it('writes a theme colour as a:schemeClr and keeps the reference', async () => {
		const fill: FillInput = {
			type: 'solid',
			color: '#ED7D31',
			themeColorRef: { scheme: 'accent2' },
		};
		const { cells, loaded } = await roundTrip(withFill(fill));
		expect(tcPr(cells[0])).toMatch(/<a:solidFill><a:schemeClr val="accent2"/u);
		expect(loaded[0]?.style?.backgroundColorRef).toEqual({ scheme: 'accent2' });
	});

	it('writes a linear gradient with its stops and angle', async () => {
		const fill: FillInput = {
			type: 'gradient',
			angle: 45,
			stops: [
				{ color: '#FFFFFF', position: 0 },
				{ color: '#1F4E79', position: 100, opacity: 0.5 },
			],
		};
		const { cells, loaded } = await roundTrip(withFill(fill));
		const xml = tcPr(cells[0]);
		expect(xml).toMatch(/<a:gs pos="0"><a:srgbClr val="FFFFFF"/u);
		expect(xml).toMatch(/<a:gs pos="100000"><a:srgbClr val="1F4E79"><a:alpha val="50000"/u);
		expect(xml).toMatch(/<a:lin ang="2700000"/u);
		expect(loaded[0]?.style).toMatchObject({ fillMode: 'gradient', gradientFillAngle: 45 });
		expect(loaded[0]?.style?.gradientFillStops?.map((s) => s.color)).toEqual([
			'#FFFFFF',
			'#1F4E79',
		]);
	});

	it('writes the documented default angle 0 (left to right) when angle is omitted', async () => {
		const fill: FillInput = {
			type: 'gradient',
			stops: [
				{ color: '#000000', position: 0 },
				{ color: '#FFFFFF', position: 100 },
			],
		};
		const { cells } = await roundTrip(withFill(fill));
		expect(tcPr(cells[0])).toMatch(/<a:lin ang="0"/u);
	});

	it('writes a path gradient with its focus', async () => {
		const fill: FillInput = {
			type: 'gradient',
			path: 'rect',
			focus: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 },
			stops: [
				{ color: '#FFFFFF', position: 0 },
				{ color: '#000000', position: 100 },
			],
		};
		const { cells, loaded } = await roundTrip(withFill(fill));
		expect(tcPr(cells[0])).toMatch(
			/<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/u,
		);
		expect(loaded[0]?.style).toMatchObject({
			gradientFillType: 'radial',
			gradientFillPathType: 'rect',
		});
	});

	it('writes a pattern with its colours', async () => {
		const fill: FillInput = {
			type: 'pattern',
			preset: 'ltDnDiag',
			foreground: '#C00000',
			background: '#FFFFFF',
		};
		const { cells, loaded } = await roundTrip(withFill(fill));
		expect(tcPr(cells[0])).toMatch(
			/<a:pattFill prst="ltDnDiag"><a:fgClr><a:srgbClr val="C00000"[\s\S]*<a:bgClr><a:srgbClr val="FFFFFF"/u,
		);
		expect(loaded[0]?.style).toMatchObject({
			fillMode: 'pattern',
			patternFillPreset: 'ltDnDiag',
			patternFillForeground: '#C00000',
			patternFillBackground: '#FFFFFF',
		});
	});

	it("writes an explicit a:noFill for type 'none', over the default table style", async () => {
		const { cells, loaded } = await roundTrip(withFill({ type: 'none' }));
		expect(tcPr(cells[0])).toMatch(/<a:noFill/u);
		expect(tcPr(cells[1])).not.toMatch(/<a:noFill/u);
		expect(loaded[0]?.style?.fillMode).toBe('none');
	});

	it('keeps the text style beside the fill', async () => {
		const input: TableInput = {
			style: 'none',
			rows: [
				{
					cells: [
						{
							text: 'both',
							style: { bold: true, color: '#FFFFFF' },
							fill: { type: 'solid', color: '#203864' },
						},
					],
				},
			],
		};
		const el = createTableElement(input);
		expect(el.tableData.rows[0]?.cells[0]?.style).toMatchObject({
			bold: true,
			color: '#FFFFFF',
			fillMode: 'solid',
			backgroundColor: '#203864',
		});
		const { cells, loaded } = await roundTrip(input);
		expect(cells[0]).toMatch(/<a:rPr[^>]*\bb="1"/u);
		expect(loaded[0]?.style).toMatchObject({ bold: true, backgroundColor: '#203864' });
	});

	it('fills a merged cell but not the slots its merge covers', async () => {
		const { cells, loaded } = await roundTrip({
			fillMergedCells: true,
			rows: [
				{ cells: [{ text: 'A', gridSpan: 2, fill: { type: 'solid', color: '#00B050' } }] },
				{ cells: [{ text: 'B' }, { text: 'C' }] },
			],
		});
		expect(tcPr(cells[0])).toMatch(/<a:srgbClr val="00B050"/u);
		expect(tcPr(cells[1])).not.toMatch(/00B050/u);
		expect(loaded[1]?.hMerge).toBe(true);
		expect(loaded[1]?.style?.backgroundColor).toBeUndefined();
	});

	it('refuses a fill a cell cannot carry', () => {
		const refused: FillInput[] = [
			{ type: 'image', url: 'data:image/png;base64,AAAA' },
			{ type: 'solid', color: '#FF0000', opacity: 0.5 },
			{ type: 'pattern', preset: '' },
			{ type: 'gradient', stops: [] },
			{
				type: 'gradient',
				tile: { l: 0, t: 0, r: 0, b: 0 },
				stops: [{ color: '#000000', position: 0 }],
			},
			{ type: 'gradient', scaled: false, stops: [{ color: '#000000', position: 0 }] },
			{ type: 'gradient', rotWithShape: true, stops: [{ color: '#000000', position: 0 }] },
			{
				type: 'gradient',
				focus: { l: 0, t: 0, r: 0, b: 0 },
				stops: [{ color: '#000000', position: 0 }],
			},
		];
		for (const fill of refused) {
			expect(() => createTableElement(withFill(fill)), JSON.stringify(fill)).toThrow(RangeError);
		}
		expect(() =>
			createTableElement(withFill({ type: 'solid', color: '#FF0000', opacity: 1 })),
		).not.toThrow();
	});
});

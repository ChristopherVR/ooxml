import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { ChartObject, DrawingAnchor } from '../model';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { createEditSession } from './session';

const fixture = (name: string) =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name)));
const anchor: DrawingAnchor = {
	from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
	to: { row: 10, col: 6, rowOffset: 0, colOffset: 0 },
};
const PNG = Uint8Array.from(
	atob(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
	),
	(c) => c.charCodeAt(0),
);
const partText = async (bytes: Uint8Array, name: string) =>
	(await JSZip.loadAsync(bytes)).file(name)?.async('text');

describe('charts and pictures', () => {
	it('saves editable WordArt, then patches the text and outline after reopening', async () => {
		const workbook = createWorkbook();
		const value = {
			text: 'Sales',
			style: {
				fill: '1769A5',
				outline: '103D63',
				gradient: { end: 'B7DDF5', angle: 16200000 },
				warp: { preset: 'textArchUp', guides: [{ name: 'adj', formula: 'val 10800000' }] },
			},
		};
		const position: DrawingAnchor = { from: anchor.from, ext: { cx: 2100000, cy: 600000 } };
		createEditSession(workbook).addWordArt(0, position, value);
		const original = await saveXlsx(workbook);
		expect(await partText(original, 'xl/drawings/drawing1.xml')).toContain('a:gradFill');
		const loaded = await loadXlsx(original);
		expect(loaded.sheets[0]?.drawings[0]).toMatchObject({
			shape: { text: { text: 'Sales' }, wordArt: value.style },
		});
		const edit = createEditSession(loaded);
		edit.editWordArt(0, 0, {
			text: 'Revenue & costs',
			style: {
				fill: 'FFFFFF',
				outline: '222222',
				shadow: '45460D',
				warp: { preset: 'textWave1', guides: [] },
			},
		});
		const saved = await saveXlsx(loaded);
		expect(await partText(saved, 'xl/drawings/drawing1.xml')).toContain('Revenue &amp; costs');
		expect((await loadXlsx(saved)).sheets[0]?.drawings[0]).toMatchObject({
			shape: {
				text: { text: 'Revenue & costs' },
				wordArt: {
					fill: 'FFFFFF',
					outline: '222222',
					shadow: '45460D',
					warp: { preset: 'textWave1', guides: [] },
				},
			},
		});
		edit.undo();
		expect(loaded.sheets[0]?.drawings[0]).toMatchObject({ shape: { text: { text: 'Sales' } } });
		edit.editWordArt(0, 0, { text: 'Sales', style: { fill: '1769A5' } });
		expect(await partText(await saveXlsx(loaded), 'xl/drawings/drawing1.xml')).not.toContain(
			'prstTxWarp',
		);
	});

	it('adds a chart and a picture to a new workbook that survive a round trip', async () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setRangeValues(0, { row: 0, col: 0 }, [
			['a', 1],
			['b', 2],
		]);
		const index = s.addChart(0, {
			anchor,
			chartType: 'column',
			title: 'Sales',
			showLegend: true,
			series: [
				{
					valuesRef: 'Sheet1!$B$1:$B$2',
					categoriesRef: 'Sheet1!$A$1:$A$2',
					categories: [],
					values: [],
				},
			],
		});
		expect(index).toBe(0);
		expect(s.undoLabel()).toBe('Insert chart');
		const svg = new TextEncoder().encode(
			'<svg xmlns="http://www.w3.org/2000/svg"><path d="m0 0 1 1"/></svg>',
		);
		const image = s.addImage(0, PNG, 'image/png', anchor, 'Logo', svg);
		expect(image.partName).toBe('xl/media/image1.png');
		expect(image.svgPartName).toBe('xl/media/image1.svg');
		const bytes = await saveXlsx(wb);
		const back = await loadXlsx(bytes);
		const kinds = back.sheets[0]?.drawings.map((d) => d.kind);
		expect(kinds).toEqual(['chart', 'image']);
		expect(back.sheets[0]?.drawings[1]).toMatchObject({ svgPartName: 'xl/media/image1.svg' });
		const zip = await JSZip.loadAsync(bytes);
		expect(await zip.file('xl/media/image1.png')?.async('uint8array')).toEqual(PNG);
		expect(await zip.file('xl/media/image1.svg')?.async('uint8array')).toEqual(svg);
		expect(await partText(bytes, 'xl/drawings/drawing1.xml')).toContain('asvg:svgBlip');
		s.undo();
		expect(wb.sheets[0]?.drawings.length).toBe(1);
	});

	it('patches a loaded chart part in place, keeping unmodelled detail', async () => {
		const source = fixture('excel-features.xlsx');
		const wb = await loadXlsx(source);
		const sheetIndex = wb.sheets.findIndex((sh) => sh.drawings.some((d) => d.kind === 'chart'));
		const sheet = wb.sheets[sheetIndex];
		if (!sheet) throw new Error('no chart in fixture');
		const index = sheet.drawings.findIndex((d) => d.kind === 'chart');
		const chart = sheet.drawings[index] as ChartObject;
		const part = chart.partName ?? '';
		const s = createEditSession(wb, { recalc: false });
		s.updateChart(sheetIndex, index, { title: 'Renamed chart', showLegend: false });
		expect((sheet.drawings[index] as ChartObject).partName).toBe(part);
		const bytes = await saveXlsx(wb);
		const before = (await partText(source, part)) ?? '';
		const after = (await partText(bytes, part)) ?? '';
		expect(after).toContain('Renamed chart');
		expect(after).not.toContain('<c:legend>');
		// Styling the model does not carry (text properties, axis formatting) is still there.
		const plot = (xml: string) => /<c:plotArea>[\s\S]*<\/c:plotArea>/.exec(xml)?.[0];
		expect(plot(before)).toBeDefined();
		expect(plot(after)).toBe(plot(before));
		const back = await loadXlsx(bytes);
		const reread = back.sheets[sheetIndex]?.drawings[index] as ChartObject;
		expect(reread.title).toBe('Renamed chart');
		expect(reread.showLegend).toBe(false);
		// A type change regenerates the part under the same name.
		s.updateChart(sheetIndex, index, { chartType: 'line' });
		const line = await loadXlsx(await saveXlsx(wb));
		expect((line.sheets[sheetIndex]!.drawings[index] as ChartObject).chartType).toBe('line');
	});
});

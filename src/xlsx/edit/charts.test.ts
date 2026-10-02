import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { ChartObject, DrawingAnchor } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/index.js';
import { createEditSession } from './session.js';

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
		const image = s.addImage(0, PNG, 'image/png', anchor, 'Logo');
		expect(image.partName).toBe('xl/media/image1.png');
		const bytes = await saveXlsx(wb);
		const back = await loadXlsx(bytes);
		const kinds = back.sheets[0]?.drawings.map((d) => d.kind);
		expect(kinds).toEqual(['chart', 'image']);
		const zip = await JSZip.loadAsync(bytes);
		expect(await zip.file('xl/media/image1.png')?.async('uint8array')).toEqual(PNG);
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
		expect((line.sheets[sheetIndex]?.drawings[index] as ChartObject).chartType).toBe('line');
	});
});

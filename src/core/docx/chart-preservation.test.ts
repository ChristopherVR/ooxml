import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { setChartTitle } from '../chart';
import { chartsIn, documentDrawingBlocks, editDocxChart, loadDocx, saveDocx } from './index';
import { isDocxChartEditable } from './chart-edit';

async function fixture(insert: (xml: string) => string) {
	const zip = await JSZip.loadAsync(
		readFileSync(path.join(import.meta.dirname, '__fixtures__/chart.docx')),
	);
	const part = 'word/charts/chart1.xml';
	const source = insert(await zip.file(part)!.async('string'));
	zip.file(part, source);
	return { bytes: await zip.generateAsync({ type: 'uint8array' }), part, source };
}

describe('Word chart source preservation', () => {
	it('keeps protection through a supported title edit and package reload', async () => {
		const { bytes, part } = await fixture((xml) =>
			xml.replace('<c:chart>', '<c:protection><c:data val="1"/></c:protection><c:chart>'),
		);
		const loaded = await loadDocx(bytes);
		const chart = chartsIn(documentDrawingBlocks(loaded.model))[0]!;
		expect(isDocxChartEditable(chart)).toBe(true);
		const edited = editDocxChart(chart, (space) => setChartTitle(space, 'Kept'));
		Object.assign(chart, edited);
		const output = await saveDocx(loaded.model);
		const zip = await JSZip.loadAsync(output);
		expect(await zip.file(part)!.async('string')).toContain('<c:protection>');
		expect(chartsIn(documentDrawingBlocks((await loadDocx(output)).model))[0]!.title).toBe('Kept');
	});
	it.each([
		['unknown series child', (xml: string) => xml.replace('</c:ser>', '<c:unknown/></c:ser>')],
		[
			'invalid number',
			(xml: string) =>
				xml.replace('<c:plotArea>', '<c:view3D><c:rotX val="bad"/></c:view3D><c:plotArea>'),
		],
	] as const)('refuses %s but preserves an untouched part', async (_, insert) => {
		const { bytes, part, source } = await fixture(insert);
		const loaded = await loadDocx(bytes);
		const chart = chartsIn(documentDrawingBlocks(loaded.model))[0]!;
		expect(isDocxChartEditable(chart)).toBe(false);
		expect(() => editDocxChart(chart, (space) => setChartTitle(space, 'Unsafe'))).toThrow();
		const saved = await JSZip.loadAsync(await saveDocx(loaded.model));
		expect(await saved.file(part)!.async('string')).toBe(source);
		// A direct model mutation cannot bypass the save-time source check.
		chart.chartSpace = setChartTitle(chart.chartSpace!, 'Unsafe');
		await expect(saveDocx(loaded.model)).rejects.toThrow('cannot be preserved');
	});
});

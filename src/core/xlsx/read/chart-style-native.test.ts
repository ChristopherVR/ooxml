import JSZip from 'jszip';
import { expect, it } from 'vitest';
import native from '../../chart/excel-chart-styles.json';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { loadXlsx } from './index';
import { saveXlsx } from '../write/index';
import { SourceIndex } from './package';
import { readChartStylePart } from './chart-colors';

for (const sample of native.cases)
	it(`imports and preserves native chart style ${sample.style} while editing the title`, async () => {
		const book = createWorkbook();
		createEditSession(book).addChart(0, {
			chartType: 'column',
			showLegend: true,
			series: [],
			anchor: {
				from: { col: 4, row: 1, colOffset: 0, rowOffset: 0 },
				ext: { cx: 6000000, cy: 4000000 },
			},
		});
		const zip = await JSZip.loadAsync(await saveXlsx(book));
		for (const [name, xml] of Object.entries(sample.parts)) zip.file(name, xml);
		const imported = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const chart = imported.sheets[0]!.drawings[0]!;
		expect(chart.kind).toBe('chart');
		if (chart.kind !== 'chart') throw Error('Native chart missing');
		expect(chart.styleDefinition?.id).toBe(sample.style);
		createEditSession(imported).updateChart(0, 0, { title: 'Updated title' });
		const saved = await saveXlsx(imported);
		const afterZip = await JSZip.loadAsync(saved);
		expect(await afterZip.file('xl/charts/style1.xml')!.async('string')).toBe(
			sample.parts['xl/charts/style1.xml'],
		);
		const reloaded = await loadXlsx(saved);
		const after = reloaded.sheets[0]!.drawings[0]!;
		if (after.kind !== 'chart') throw Error('Saved chart missing');
		expect(after.styleDefinition).toEqual(chart.styleDefinition);
		expect(after.title).toBe('Updated title');
	});

it('reports an unreadable style separately from its readable chart', () => {
	const sample = native.cases[0]!;
	for (const xml of ['<broken>', '<chartStyle xmlns="urn:other"/>']) {
		const parts = new Map(
			Object.entries(sample.parts).map(([name, text]) => [name, new TextEncoder().encode(text)]),
		);
		parts.set('xl/charts/style1.xml', new TextEncoder().encode(xml));
		const warnings: string[] = [];
		expect(
			readChartStylePart(new SourceIndex(parts), 'xl/charts/chart1.xml', (message) =>
				warnings.push(message),
			),
		).toBeUndefined();
		expect(warnings).toEqual([
			'Chart style for xl/charts/chart1.xml could not be read; its part is kept but formatting is not shown.',
		]);
	}
});

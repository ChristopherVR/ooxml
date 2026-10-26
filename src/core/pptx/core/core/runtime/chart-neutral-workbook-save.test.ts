import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	addChartSeries,
	createChartWorkbook,
	parseChartSpace,
	writeChartSpace,
} from '../../../../chart';
import { columnChart } from '../../../../chart/edit/test-support/column-chart';
import { loadXlsx } from '../../../../xlsx/read';
import { getCell } from '../../../../xlsx/cells';
import type { PptxChartData } from '../../types';
import { prepareNeutralChartWorkbook } from './chart-neutral-workbook-save';

describe('PowerPoint structural workbook save', () => {
	it('rebases added series into a fresh worksheet and leaves original cells byte-for-byte', async () => {
		const bytes = await createChartWorkbook({
			categories: ['A', 'B'],
			series: [{ name: 'Original', values: [9, 8] }],
		});
		const oldZip = await JSZip.loadAsync(bytes);
		const original = await oldZip.file('xl/worksheets/sheet1.xml')!.async('string');
		const zip = new JSZip();
		zip.file('ppt/embeddings/chart.xlsx', bytes);
		const model = addChartSeries(columnChart(), { name: 'Added', values: [41, 42] });
		const chart = {
			chartType: 'bar',
			categories: ['Q1', 'Q2'],
			series: [{}, {}, {}],
			chartPartPath: 'ppt/charts/chart1.xml',
			externalData: { targetPath: '../embeddings/chart.xlsx' },
		} as PptxChartData;
		const result = await prepareNeutralChartWorkbook(
			zip,
			chart,
			{ xml: writeChartSpace(model), workbookModel: model, workbookWrites: [] },
			() => 'ppt/embeddings/chart.xlsx',
		);
		const saved = await zip.file('ppt/embeddings/chart.xlsx')!.async('uint8array');
		const outputZip = await JSZip.loadAsync(saved);
		expect(await outputZip.file('xl/worksheets/sheet1.xml')!.async('string')).toBe(original);
		const book = await loadXlsx(saved);
		expect(book.sheets).toHaveLength(2);
		expect(getCell(book.sheets[1]!, 1, 5)?.value).toBe(41);
		const series = parseChartSpace(result.xml).chartSpace.plotArea.groups[0]!.series;
		expect(series[2]!.values!.formula).toBe('ChartData1!$F$2:$F$3');
		expect(series[2]!.tx!.reference!.formula).toBe('ChartData1!$F$1');
		expect(chart.series[2]!.idx).toBe(series[2]!.index);
		expect(result.workbookWrites).toEqual([]);
	});
	it('refuses external workbooks without changing any package bytes', async () => {
		const zip = new JSZip();
		zip.file('keep', 'original');
		const model = columnChart();
		const chart = {
			series: [],
			externalData: { isExternal: true, targetPath: 'remote.xlsx' },
			chartPartPath: 'ppt/charts/chart1.xml',
		} as unknown as PptxChartData;
		await expect(
			prepareNeutralChartWorkbook(
				zip,
				chart,
				{ xml: writeChartSpace(model), workbookModel: model, workbookWrites: [] },
				() => 'remote.xlsx',
			),
		).rejects.toThrow('embedded');
		expect(Object.keys(zip.files)).toEqual(['keep']);
		expect(await zip.file('keep')!.async('string')).toBe('original');
	});
});

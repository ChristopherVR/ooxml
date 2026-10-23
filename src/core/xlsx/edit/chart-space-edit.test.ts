// Excel charts edited over their ChartSpace (chart step C2): intents and summary patches change the
// model the part is saved from, the summary follows the model, and untouched parts stay byte for
// byte.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { ChartObject, Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from './session';

async function book(): Promise<Workbook> {
	const workbook = createWorkbook();
	workbook.sheets[0]!.name = 'Data';
	const session = createEditSession(workbook);
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Month', 'Sales'],
		['Jan', 10],
		['Feb', 20],
	]);
	session.addChart(0, {
		anchor: { from: { row: 0, col: 4, rowOffset: 0, colOffset: 0 }, ext: { cx: 4e6, cy: 3e6 } },
		chartType: 'column',
		title: 'Sales',
		showLegend: true,
		series: [
			{
				name: 'Sales',
				nameRef: 'Data!$B$1',
				valuesRef: 'Data!$B$2:$B$3',
				categoriesRef: 'Data!$A$2:$A$3',
				values: [10, 20],
				categories: ['Jan', 'Feb'],
			},
		],
	});
	return loadXlsx(await saveXlsx(workbook));
}

const chartOf = (workbook: Workbook) => workbook.sheets[0]!.drawings[0] as ChartObject;
const part = async (workbook: Workbook) =>
	(await JSZip.loadAsync(await saveXlsx(workbook))).file('xl/charts/chart1.xml')!.async('string');

describe('xlsx charts over ChartSpace', () => {
	it('keeps the model of a read chart and saves an unedited part byte for byte', async () => {
		const workbook = await book();
		const source = new TextDecoder().decode(workbook.source!.parts.get('xl/charts/chart1.xml'));
		expect(chartOf(workbook).chartSpace?.plotArea.groups[0]?.series).toHaveLength(1);
		expect(await part(workbook)).toBe(source);
	});

	it('applies an intent, derives the summary from the model and undoes it', async () => {
		const workbook = await book();
		const session = createEditSession(workbook);
		const before = structuredClone(chartOf(workbook));
		session.editChart(0, 0, { op: 'axis', axisIndex: 1, edit: { max: 40, majorGridlines: false } });
		expect(chartOf(workbook).axes?.[1]).toMatchObject({
			scaling: { max: 40 },
			majorGridlines: false,
		});
		session.editChart(0, 0, { op: 'dataLabels', edit: { showValue: true } });
		expect(chartOf(workbook).series[0]?.dataLabels).toBeDefined();
		session.editChart(0, 0, { op: 'colorPalette', from: 10, to: 12 });
		expect(chartOf(workbook).colorPalette).toBe(12);
		const reloaded = chartOf(await loadXlsx(await saveXlsx(workbook)));
		expect(reloaded.axes?.[1]).toMatchObject({ scaling: { max: 40 }, majorGridlines: false });
		expect(reloaded.colorPalette).toBe(12);
		session.undo();
		session.undo();
		session.undo();
		expect(chartOf(workbook)).toEqual(before);
	});

	it('maps summary patches onto the model and refuses intents the model cannot take', async () => {
		const workbook = await book();
		const session = createEditSession(workbook);
		session.updateChart(0, 0, { title: 'Revenue', legendPosition: 'b', barGapWidth: 60 });
		const space = chartOf(workbook).chartSpace!;
		expect(space.title?.text).toBe('Revenue');
		expect(space.legend?.position).toBe('b');
		expect(space.plotArea.groups[0]?.gapWidth).toBe(60);
		expect(chartOf(workbook)).toMatchObject({ title: 'Revenue', barGapWidth: 60 });
		expect(() =>
			session.editChart(0, 0, { op: 'groupOptions', seriesIndex: 0, edit: { holeSize: 50 } }),
		).toThrow(RangeError);
		expect(session.undoLabel()).toBe('Edit chart');
	});

	it('moves the model references with inserted rows', async () => {
		const workbook = await book();
		createEditSession(workbook).insertRows(0, 0, 1);
		const series = chartOf(workbook).chartSpace!.plotArea.groups[0]!.series[0]!;
		expect(series.values?.formula).toBe('Data!$B$3:$B$4');
		expect(series.tx?.reference?.formula).toBe('Data!$B$2');
		const saved = chartOf(await loadXlsx(await saveXlsx(workbook)));
		expect(saved.series[0]?.valuesRef).toBe('Data!$B$3:$B$4');
	});

	it('refreshes a new chart model from its cells on save', async () => {
		const workbook = createWorkbook();
		workbook.sheets[0]!.name = 'Data';
		const session = createEditSession(workbook);
		session.setRangeValues(0, { row: 0, col: 0 }, [['Sales'], [1], [2]]);
		session.addChart(0, {
			anchor: { from: { row: 0, col: 4, rowOffset: 0, colOffset: 0 } },
			chartType: 'line',
			showLegend: false,
			series: [
				{ nameRef: 'Data!$A$1', valuesRef: 'Data!$A$2:$A$3', values: [1, 2], categories: [] },
			],
		});
		session.setRangeValues(0, { row: 0, col: 0 }, [['Revenue'], [5], [7]]);
		const saved = chartOf(await loadXlsx(await saveXlsx(workbook)));
		expect(saved.series[0]).toMatchObject({ name: 'Revenue', values: [5, 7] });
	});
});

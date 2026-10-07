import JSZip from 'jszip';
import { expect, it } from 'vitest';
import native from '../__fixtures__/excel-gradient-presets.json';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { loadXlsx } from '../read';
import { saveXlsx } from '../write';
import type { ChartObject } from '../model';
import { chartSeriesGradientPatch } from './chart-series-gradient';
import {
	officeGradientPresetFill,
	officeGradientPresetId,
	OFFICE_GRADIENT_PRESETS,
} from '../../diagram/gradient-presets';

async function nativeBook(xml: string) {
	const book = createWorkbook();
	createEditSession(book).addChart(0, {
		chartType: 'column',
		series: [],
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	});
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	zip.file('xl/charts/chart1.xml', xml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}
const chartOf = (book: ReturnType<typeof createWorkbook>) =>
	book.sheets[0]!.drawings[0] as ChartObject;

for (const sample of native.cases)
	it(`authors native Office gradient preset ${sample.id} with one undo and round trip`, async () => {
		const book = await nativeBook(sample.chartXml);
		const chart = chartOf(book);
		const source = chart.series[0]!.fill!;
		if (source.kind !== 'gradient') throw new Error('Expected native gradient');
		expect(sample.preset).toBe(sample.id);
		expect(officeGradientPresetFill(sample.id)).toMatchObject({
			angle: sample.angle,
			scaled: source.scaled,
			stops: source.stops,
		});
		expect(officeGradientPresetId(source)).toBe(sample.id);
		chart.series.push(structuredClone(chart.series[0]!));
		const untouched = structuredClone(chart.series[1]);
		const previous = officeGradientPresetFill(sample.id === 1 ? 2 : 1);
		chart.series[0]!.fill = previous;
		chart.series[0]!.pointFills = { 0: { kind: 'none' } };
		const session = createEditSession(book);
		const edit = chartSeriesGradientPatch(chart, 0, { kind: 'preset', id: sample.id })!;
		expect(edit.stopIndex).toBe(0);
		session.updateChart(0, 0, edit.patch);
		expect(chartOf(book).series[0]!.fill).toMatchObject({
			stops: source.stops,
			angle: sample.angle,
			scaled: true,
		});
		expect(chartOf(book).series[0]!.pointFills).toBeUndefined();
		expect(chartOf(book).series[1]).toEqual(untouched);
		expect(
			chartSeriesGradientPatch(chartOf(book), 0, { kind: 'preset', id: sample.id }),
		).toBeUndefined();
		const roundtrip = chartOf(await loadXlsx(await saveXlsx(book)));
		expect(roundtrip.series[0]!.fill).toMatchObject({
			stops: source.stops,
			angle: sample.angle,
			scaled: true,
		});
		expect(officeGradientPresetId(roundtrip.series[0]!.fill!)).toBe(sample.id);
		session.undo();
		expect(chartOf(book).series[0]!.fill).toEqual(previous);
		expect(chartOf(book).series[0]!.pointFills).toEqual({ 0: { kind: 'none' } });
		session.redo();
		expect(officeGradientPresetId(chartOf(book).series[0]!.fill!)).toBe(sample.id);
	});

it('guards invalid IDs and keeps customized paints and factory instances independent', () => {
	expect(OFFICE_GRADIENT_PRESETS).toHaveLength(24);
	for (const id of [-1, 0, 25, 1.5, NaN])
		expect(() => officeGradientPresetFill(id)).toThrow(RangeError);
	const first = officeGradientPresetFill(5),
		second = officeGradientPresetFill(5);
	first.stops[0]!.color.value = '123456';
	expect(officeGradientPresetId(first)).toBeUndefined();
	expect(officeGradientPresetId(second)).toBe(5);
	second.angle = 45;
	expect(officeGradientPresetId(second)).toBeUndefined();
});

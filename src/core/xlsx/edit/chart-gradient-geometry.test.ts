import { expect, it } from 'vitest';
import native from '../../chart/__fixtures__/native-gradient-path-profiles.json';
import series from '../../chart/__fixtures__/native-gradient-series-path-profiles.json';
import { parseXml, NS } from '../../xml';
import { parseDrawingFill } from '../../drawingml/drawing-fill';
import { drawingFillXml } from '../../drawingml/write-fill';
import { rectGradientDirection } from '../../drawingml/gradient-geometry';
import { officeGradientPresetFill } from '../../drawingml/gradient-presets';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { saveXlsx } from '../write';
import { loadXlsx } from '../read';
import { chartSeriesGradientPatch } from './chart-series-gradient';

for (const sample of [...native.cases, ...series.cases])
	it(`authors native target and tile geometry for ${sample.name}`, async () => {
		const expected = parseDrawingFill(
			parseXml(`<a:spPr xmlns:a="${NS.a}">${sample.fillXml}</a:spPr>`).documentElement,
		)!;
		if (expected.kind !== 'gradient') throw new Error('Expected native gradient');
		const type = expected.path;
		if (type !== 'rect' && type !== 'circle' && type !== 'shape')
			throw new Error('Expected native path type');
		const direction = rectGradientDirection(expected.fillToRect)!;
		const book = createWorkbook(),
			session = createEditSession(book);
		const original = officeGradientPresetFill(1);
		original.sourceXml = `<a:gradFill xmlns:a="${NS.a}" flip="xy" rotWithShape="0"><a:extLst/></a:gradFill>`;
		session.addChart(0, {
			chartType: 'column',
			showLegend: false,
			series: [{ name: 'Sales', categories: ['A'], values: [1], fill: original }],
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		});
		const chart = book.sheets[0]!.drawings[0]!;
		if (chart.kind !== 'chart') throw new Error('Expected chart');
		const originalChart = structuredClone(chart);
		const edit = chartSeriesGradientPatch(chart, 0, { kind: 'geometry', type, direction })!;
		expect(chart.series[0]!.fill).toEqual(original);
		session.updateChart(0, 0, edit.patch);
		const updated = book.sheets[0]!.drawings[0]!;
		if (updated.kind !== 'chart') throw new Error('Expected chart');
		expect(updated.series[0]!.fill).toMatchObject({
			path: expected.path,
			fillToRect: expected.fillToRect,
			tileRect: expected.tileRect,
			stops: original.stops,
		});
		expect(
			chartSeriesGradientPatch(updated, 0, { kind: 'geometry', type, direction }),
		).toBeUndefined();
		const roundtrip = (await loadXlsx(await saveXlsx(book))).sheets[0]!.drawings[0]!;
		if (roundtrip.kind !== 'chart') throw new Error('Expected chart');
		expect(roundtrip.series[0]!.fill).toMatchObject({
			path: expected.path,
			fillToRect: expected.fillToRect,
			tileRect: expected.tileRect,
			stops: original.stops,
		});
		const xml = drawingFillXml(roundtrip.series[0]!.fill!)!;
		expect(xml).toContain('flip="xy"');
		expect(xml).toContain('rotWithShape="0"');
		expect(xml).toContain('extLst');
		const rectangularChart = structuredClone(updated);
		const linear = chartSeriesGradientPatch(updated, 0, { kind: 'geometry', type: 'linear' })!;
		session.updateChart(0, 0, linear.patch);
		const changed = book.sheets[0]!.drawings[0]!;
		if (changed.kind !== 'chart' || changed.series[0]!.fill?.kind !== 'gradient')
			throw new Error('Expected gradient');
		expect(changed.series[0]!.fill.path).toBeUndefined();
		expect(changed.series[0]!.fill.tileRect).toEqual({ l: 0, t: 0, r: 0, b: 0 });
		expect(changed.series[0]!.fill.stops).toEqual(original.stops);
		expect(
			chartSeriesGradientPatch(changed, 0, { kind: 'geometry', type: 'linear' }),
		).toBeUndefined();
		session.undo();
		expect(book.sheets[0]!.drawings[0]).toEqual(rectangularChart);
		session.undo();
		expect(book.sheets[0]!.drawings[0]).toEqual(originalChart);
		session.redo();
		expect(book.sheets[0]!.drawings[0]).toEqual(rectangularChart);
	});

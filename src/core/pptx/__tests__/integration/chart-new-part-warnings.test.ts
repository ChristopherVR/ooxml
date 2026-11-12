/**
 * What a NEW chart's save reports: nothing for a part the shared chart core writes in full,
 * `CHART_FIELDS_NOT_WRITTEN` for model fields no new part holds, and
 * `CHART_WRITTEN_BY_LEGACY_GENERATOR` when the legacy generator must write the part.
 */
import { describe, expect, it } from 'vitest';

import { setChartAxis } from '../../core/builders/sdk/chart-operations';
import { newChartDeck } from '../helpers/new-chart-deck';

describe('new chart save warnings', () => {
	it('reports nothing for a chart the shared core writes in full', async () => {
		const { save } = await newChartDeck('bar', { title: 'Sales' });
		expect((await save()).warnings).toStrictEqual([]);
	});

	it('reports an axis edit on a type without axes, and writes the rest', async () => {
		const { element, save } = await newChartDeck('pie');
		setChartAxis(element, 'valAx', { title: 'Nowhere' });
		const { xml, chart, warnings } = await save();
		expect(xml).not.toContain('<c:valAx>');
		expect(chart.chartType).toBe('pie');
		expect(warnings).toMatchObject([
			{ code: 'CHART_FIELDS_NOT_WRITTEN', scope: 'save', elementId: element.id },
		]);
		expect(warnings[0]!.message).toContain('axes.0.valAx');
	});

	it('reports user shapes, which a new part has no drawing for', async () => {
		const { element, save } = await newChartDeck('bar');
		element.chartData!.userShapesXml = { 'c:userShapes': {} };
		const { warnings } = await save();
		expect(warnings.map((warning) => warning.code)).toStrictEqual(['CHART_FIELDS_NOT_WRITTEN']);
		expect(warnings[0]!.message).toContain('userShapesXml');
	});

	it('names the cause when the legacy generator writes the part', async () => {
		const unknown = await newChartDeck('bar');
		unknown.element.chartData!.chartType = 'unknown';
		const first = await unknown.save();
		expect(first.warnings.map((warning) => warning.code)).toStrictEqual([
			'CHART_WRITTEN_BY_LEGACY_GENERATOR',
		]);
		expect(first.warnings[0]!.message).toContain('chartType.unknown');
		const pivot = await newChartDeck('bar');
		pivot.element.chartData!.pivotSource = { name: '[Book1]Sheet1!PivotTable1', formatId: 0 };
		expect((await pivot.save()).warnings[0]!.message).toContain('pivotSource');
	});

	it('keeps the legacy generator for values the shared core refuses', async () => {
		const { element, save } = await newChartDeck('line');
		element.chartData!.series[0]!.values = [1, Number.NaN, 3];
		const { warnings } = await save();
		expect(warnings.map((warning) => warning.code)).toStrictEqual([
			'CHART_WRITTEN_BY_LEGACY_GENERATOR',
		]);
		expect(warnings[0]!.message).toContain('finite numbers');
	});
});

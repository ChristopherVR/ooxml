import { describe, expect, it } from 'vitest';

import { parseChartSpace, writeChartSpace } from '../../../chart/index';
import type { PptxChartData } from '../types';
import { planNeutralChartSave } from './chart-neutral-save-plan';
import { saveChartPartThroughNeutral } from './chart-neutral-save';

const PART =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
	'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
	'<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
	'<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
	'<c:ser><c:idx val="3"/><c:order val="3"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>' +
	'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
	'<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>4.4000000000000004</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>' +
	'<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>' +
	'<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
	'<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea>' +
	'<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

const before = (): PptxChartData =>
	({
		chartType: 'bar',
		title: 'Sales',
		titleRuns: [{ text: 'Sales', bold: true }],
		categories: ['Q1', 'Q2'],
		series: [{ name: 'North', idx: 3, values: [4.4, 2] }],
		style: { hasTitle: true, hasLegend: true, legendPosition: 'r', legendOverlay: false },
		axes: [
			{ axisType: 'catAx', axisId: 1, orientation: 'minMax' },
			{ axisType: 'valAx', axisId: 2, orientation: 'minMax' },
		],
	}) as PptxChartData;

const edited = (edit: (chart: PptxChartData) => void): PptxChartData => {
	const chart = before();
	edit(chart);
	return chart;
};

const save = (edit: (chart: PptxChartData) => void) =>
	saveChartPartThroughNeutral(PART, before(), edited(edit));

describe('planNeutralChartSave', () => {
	it('plans nothing for an unchanged chart, and an empty style is no change', () => {
		expect(planNeutralChartSave(before(), before())).toEqual({ edits: [], seriesData: false });
		const withoutStyle = { ...before(), style: undefined };
		expect(planNeutralChartSave(withoutStyle, { ...withoutStyle, style: {} })?.edits).toEqual([]);
	});

	it('leaves edits no family claims to the object-tree writers', () => {
		expect(
			planNeutralChartSave(
				before(),
				edited((c) => (c.chartType = 'line')),
			),
		).toBeUndefined();
		const runs = edited((c) => (c.titleRuns = [{ text: 'Sales', italic: true }]));
		expect(planNeutralChartSave(before(), runs)).toBeUndefined();
		const entries = edited((c) => (c.style!.legendEntries = [{ index: 0, deleted: true }]));
		expect(planNeutralChartSave(before(), entries)).toBeUndefined();
	});

	it('plans multi-run title text through the shared run-preserving policy', () => {
		const multi = { ...before(), titleRuns: [{ text: 'Sa' }, { text: 'les', bold: true }] };
		expect(planNeutralChartSave(multi, { ...multi, title: 'Revenue' })?.edits).toHaveLength(1);
	});
});

describe('saveChartPartThroughNeutral: title and legend', () => {
	it('retitles the first run, keeping its formatting and every other element as written', () => {
		const saved = save((c) => (c.title = 'Revenue'))!;
		expect(saved.xml).toBe(PART.replace('<a:t>Sales</a:t>', '<a:t>Revenue</a:t>'));
	});

	it('removes the title and sets autoTitleDeleted, for hasTitle false or an empty title', () => {
		for (const edit of [
			(c: PptxChartData) => (c.style!.hasTitle = false),
			(c: PptxChartData) => {
				c.title = '';
				delete c.style!.hasTitle;
			},
		]) {
			const { chartSpace } = parseChartSpace(save(edit)!.xml);
			expect(chartSpace.title).toBeUndefined();
			expect(chartSpace.autoTitleDeleted).toBe(true);
		}
	});

	it('moves, overlays and hides the legend', () => {
		const moved = parseChartSpace(
			save((c) => {
				c.style!.legendPosition = 't';
				c.style!.legendOverlay = true;
			})!.xml,
		).chartSpace;
		expect(moved.legend).toMatchObject({ position: 't', overlay: true });
		const hidden = parseChartSpace(save((c) => (c.style!.hasLegend = false))!.xml).chartSpace;
		expect(hidden.legend).toBeUndefined();
	});

	it('ignores a legend position without hasLegend, as the PowerPoint writer does', () => {
		const silent = { ...before(), style: { ...before().style, hasLegend: undefined } };
		const plan = planNeutralChartSave(silent, {
			...silent,
			style: { ...silent.style, legendPosition: 'b' },
		});
		expect(plan?.edits).toEqual([]);
	});

	it('declines a part the neutral parser reports unmodelled elements in', () => {
		const withBands = PART.replace(
			'<c:gapWidth val="150"/>',
			'<c:gapWidth val="150"/><c:unknownThing/>',
		);
		expect(parseChartSpace(withBands).issues.length).toBeGreaterThan(0);
		expect(
			saveChartPartThroughNeutral(
				withBands,
				before(),
				edited((c) => (c.title = 'X')),
			),
		).toBeUndefined();
	});

	it('hands back the workbook writes of every referenced range', () => {
		expect(save((c) => (c.title = 'Revenue'))!.workbookWrites).toEqual([
			{ formula: 'Sheet1!$B$1', isNumeric: false, values: ['North'] },
			{ formula: 'Sheet1!$A$2:$A$3', isNumeric: false, values: ['Q1', 'Q2'] },
			{ formula: 'Sheet1!$B$2:$B$3', isNumeric: true, values: ['4.4', '2'] },
		]);
		expect(writeChartSpace(parseChartSpace(PART).chartSpace)).toBe(PART);
	});
});

describe('saveChartPartThroughNeutral: axes', () => {
	const valueAxis = (xml: string) =>
		parseChartSpace(xml).chartSpace.plotArea.axes.find((axis) => axis.kind === 'val')!;

	it('writes bounds, units, visibility, number format and direction by axis id', () => {
		const saved = save((c) => {
			Object.assign(c.axes![1]!, {
				min: 0,
				max: 50,
				majorUnit: 10,
				deleted: true,
				orientation: 'maxMin',
				numFmt: { formatCode: '0.0%', sourceLinked: false },
			});
		})!;
		expect(valueAxis(saved.xml)).toMatchObject({
			scaling: { min: 0, max: 50, orientation: 'maxMin' },
			majorUnit: 10,
			deleted: true,
			numberFormat: { formatCode: '0.0%', sourceLinked: false },
		});
	});

	it('leaves axis edits the shared operation does not make to the object-tree writers', () => {
		const added = edited((c) => c.axes!.push({ axisType: 'valAx', axisId: 9 }));
		expect(planNeutralChartSave(before(), added)).toBeUndefined();
		// A minimum past the maximum is rejected by the shared operation.
		expect(save((c) => Object.assign(c.axes![1]!, { min: 9, max: 1 }))).toBeUndefined();
	});

	it('saves linked formats, gridlines, log scale and axis titles through the neutral model', () => {
		const result = save((chart) =>
			Object.assign(chart.axes![1]!, {
				numFmt: { formatCode: '0.00', sourceLinked: true },
				majorGridlines: true,
				titleText: 'Revenue',
				logScale: true,
				logBase: 2,
				displayUnits: 'millions',
			}),
		)!;
		expect(valueAxis(result.xml)).toMatchObject({
			numberFormat: { formatCode: '0.00', sourceLinked: true },
			majorGridlines: true,
			title: { text: 'Revenue' },
			scaling: { logBase: 2 },
			displayUnits: { builtInUnit: 'millions' },
		});
		expect(result.xml).toContain('<c:v>4.4000000000000004</c:v>');
	});
});

describe('saveChartPartThroughNeutral: series data', () => {
	const series = (xml: string) => parseChartSpace(xml).chartSpace.plotArea.groups[0]!.series[0]!;

	it('renames through the referenced cell, keeping the formula, and writes the workbook back', () => {
		const saved = save((c) => (c.series[0]!.name = 'South'))!;
		expect(series(saved.xml).tx?.reference).toMatchObject({
			formula: 'Sheet1!$B$1',
			cache: { points: [{ index: 0, value: 'South' }] },
		});
		expect(saved.workbookWrites[0]).toEqual({
			formula: 'Sheet1!$B$1',
			isNumeric: false,
			values: ['South'],
		});
	});

	it('writes edited values with blanks as missing points, and new text categories', () => {
		const saved = save((c) => {
			c.series[0]!.values = [7, 0];
			c.series[0]!.blanks = [false, true];
			c.categories = ['H1', 'H2'];
		})!;
		const written = series(saved.xml);
		expect(written.values?.cache).toMatchObject({
			pointCount: 2,
			points: [{ index: 0, value: '7' }],
		});
		expect(written.values?.formula).toBe('Sheet1!$B$2:$B$3');
		expect(written.categories?.cache?.points.map((p) => p.value)).toEqual(['H1', 'H2']);
		// Untouched series and caches keep the text Office wrote ("4.4000000000000004").
		expect(save((c) => (c.title = 'Revenue'))!.xml).toContain('<c:v>4.4000000000000004</c:v>');
	});

	it('leaves added series and multi-level or date categories to the object-tree writers', () => {
		const added = edited((c) => c.series.push({ name: 'New', values: [1, 2] }));
		expect(planNeutralChartSave(before(), added)).toBeUndefined();
		const levels = { ...before(), categoryLevels: [['A', 'B']] };
		expect(planNeutralChartSave(levels, { ...levels, categories: ['X', 'Y'] })).toBeUndefined();
	});
});

describe('saveChartPartThroughNeutral: labels and series paint', () => {
	it('shows labels without rewriting a series cache or its theme colour', () => {
		const source = PART.replace(
			'<c:cat>',
			'<c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr><c:cat>',
		);
		const next = edited((chart) => {
			chart.style!.hasDataLabels = true;
			chart.style!.dataLabels = { showValue: true, position: 'outEnd' };
		});
		const saved = saveChartPartThroughNeutral(source, before(), next)!;
		expect(saved).toBeDefined();
		expect(saved.xml).toContain('<a:schemeClr val="accent2"/>');
		expect(saved.xml).toContain('<c:v>4.4000000000000004</c:v>');
		expect(parseChartSpace(saved.xml).chartSpace.plotArea.groups[0]!.dataLabels).toMatchObject({
			showValue: true,
			position: 'outEnd',
		});
	});

	it('recolours a series while retaining its transparency and unrelated title runs', () => {
		const source = PART.replace(
			'<c:cat>',
			'<c:spPr><a:solidFill><a:schemeClr val="accent2"><a:alpha val="60000"/></a:schemeClr></a:solidFill></c:spPr><c:cat>',
		);
		const next = edited((chart) => {
			chart.series[0]!.color = '#FF0000';
		});
		const saved = saveChartPartThroughNeutral(source, before(), next)!;
		expect(saved).toBeDefined();
		expect(saved.xml).toContain('<a:rPr lang="en-US" b="1"/>');
		expect(saved.xml).toContain('<c:v>4.4000000000000004</c:v>');
		expect(
			parseChartSpace(saved.xml).chartSpace.plotArea.groups[0]!.series[0]!.spPr?.fill,
		).toMatchObject({
			kind: 'solid',
			color: { kind: 'srgb', value: 'FF0000', transforms: [{ name: 'alpha', value: '60000' }] },
		});
	});

	it('does not claim a value-from-cells label edit or a hidden-label policy change', () => {
		const range = edited((chart) => {
			chart.style!.hasDataLabels = true;
			chart.style!.dataLabels = {
				showValue: true,
				dataLabelsRange: { formula: 'Sheet1!$C$2:$C$3', cache: ['A', 'B'] },
			};
		});
		expect(planNeutralChartSave(before(), range)).toBeUndefined();
		expect(
			planNeutralChartSave(
				before(),
				edited((chart) => {
					chart.style!.hasDataLabels = false;
				}),
			),
		).toBeUndefined();
	});
});

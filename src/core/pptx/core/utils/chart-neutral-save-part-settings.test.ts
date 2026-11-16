import { describe, expect, it } from 'vitest';

import type { PptxChartData } from '../types';
import { planNeutralChartSave } from './chart-neutral-save-plan';
import { saveChartPartThroughNeutral } from './chart-neutral-save';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PART =
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="${C}" xmlns:r="${R}">` +
	'<c:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
	'<c:chart><c:autoTitleDeleted val="1"/><c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts><c:plotArea><c:layout/>' +
	'<c:pieChart><c:varyColors val="1"/>' +
	'<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>' +
	'<c:firstSliceAng val="0"/></c:pieChart></c:plotArea><c:plotVisOnly val="1"/></c:chart>' +
	'<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' +
	'<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>' +
	'<c:userShapes r:id="rId2"/></c:chartSpace>';

const before = (): PptxChartData =>
	({
		chartType: 'pie',
		categories: [],
		series: [{ name: '', idx: 0, values: [4] }],
		externalData: { relId: 'rId1', autoUpdate: false },
		clrMapOvr: { bg1: 'lt1' },
		printSettings: { pageSetup: {} },
		pivotFormats: { formats: [{ index: 0 }] },
		userShapes: [{ id: 'a' }],
	}) as unknown as PptxChartData;

const edited = (edit: (chart: PptxChartData) => void) => {
	const chart = before();
	edit(chart);
	return chart;
};

describe('neutral save of chart part settings', () => {
	it('writes the auto-update flag and a colour mapping', () => {
		const xml = saveChartPartThroughNeutral(
			PART,
			before(),
			edited((chart) => {
				chart.externalData!.autoUpdate = true;
				chart.clrMapOvr = {
					bg1: 'dk1',
					tx1: 'lt1',
					bg2: 'dk2',
					tx2: 'lt2',
					accent1: 'accent2',
					accent2: 'accent1',
					accent3: 'accent3',
					accent4: 'accent4',
					accent5: 'accent5',
					accent6: 'accent6',
					hlink: 'hlink',
					folHlink: 'folHlink',
				};
			}),
		)!.xml;
		expect(xml).toContain('<c:autoUpdate val="1"/>');
		expect(xml).toMatch(/<c:clrMapOvr bg1="dk1" tx1="lt1"[^>]*\/><c:chart>/u);
		const removed = saveChartPartThroughNeutral(
			PART,
			before(),
			edited((chart) => (chart.clrMapOvr = null)),
		)!.xml;
		expect(removed).not.toContain('clrMapOvr');
	});

	it('leaves a partial colour mapping to the object-tree path', () => {
		expect(
			saveChartPartThroughNeutral(
				PART,
				before(),
				edited((chart) => (chart.clrMapOvr = { bg1: 'dk1' })),
			),
		).toBeUndefined();
	});

	it('places the elements the package adapter encoded', () => {
		const after = edited((chart) => {
			chart.printSettings = { pageSetup: { orientation: 'landscape' } };
			chart.pivotFormats = null;
			chart.userShapes = [];
		});
		expect(planNeutralChartSave(before(), after)?.packageElements).toEqual([
			'printSettings',
			'pivotFormats',
			'userShapes',
		]);
		// Without them the part keeps the object-tree path.
		expect(saveChartPartThroughNeutral(PART, before(), after)).toBeUndefined();
		const xml = saveChartPartThroughNeutral(PART, before(), after, {
			elements: {
				printSettings: `<c:printSettings xmlns:c="${C}"><c:pageSetup orientation="landscape"/></c:printSettings>`,
				pivotFormats: null,
				userShapes: null,
			},
		})!.xml;
		expect(xml).toMatch(
			/<\/c:externalData><c:printSettings><c:pageSetup orientation="landscape"\/><\/c:printSettings><\/c:chartSpace>/u,
		);
		expect(xml).not.toContain('pivotFmts');
		expect(xml).not.toContain('userShapes');
	});
});

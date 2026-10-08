import { describe, expect, it } from 'vitest';
import { NS } from '../xml/index';
import type { ChartSpace } from './model';
import { parseChartSpace } from './parse-space';
import { chartNumber } from './write-util';
import { writeChartSpace } from './write-space';
import { stripDeclarations } from './xml-fragment';

const ROOT = `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">`;
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Drops the kept sources, which a re-parse adds to a model written from scratch. */
const withoutSources = (value: unknown): unknown =>
	JSON.parse(JSON.stringify(value, (key, item) => (key === 'sourceXml' ? undefined : item)));

const column: ChartSpace = {
	roundedCorners: false,
	title: { text: 'Sales & cost' },
	autoTitleDeleted: false,
	plotArea: {
		layout: {},
		groups: [
			{
				kind: 'bar',
				element: 'barChart',
				is3D: false,
				barDirection: 'col',
				grouping: 'clustered',
				varyColors: false,
				series: [
					{
						index: 0,
						order: 0,
						tx: { value: 'Sales', text: 'Sales' },
						spPr: {
							fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4', transforms: [] } },
						},
						invertIfNegative: false,
						dataPoints: [],
						categories: {
							kind: 'strRef',
							formula: "'Q&A'!$A$2:$A$3",
							cache: {
								type: 'string',
								pointCount: 2,
								points: [
									{ index: 0, value: 'North' },
									{ index: 1, value: 'South' },
								],
							},
						},
						values: {
							kind: 'numLit',
							cache: {
								type: 'number',
								formatCode: 'General',
								pointCount: 2,
								points: [{ index: 0, value: '1.50' }],
							},
						},
					},
				],
				gapWidth: 150,
				axisIds: [1, 2],
			},
		],
		axes: [
			{
				kind: 'cat',
				id: 1,
				crossAxisId: 2,
				position: 'b',
				scaling: { orientation: 'minMax' },
				majorGridlines: false,
				minorGridlines: false,
			},
			{
				kind: 'val',
				id: 2,
				crossAxisId: 1,
				position: 'l',
				scaling: { orientation: 'minMax', max: 0.0000001 },
				majorGridlines: true,
				minorGridlines: false,
				crossBetween: 'between',
			},
		],
	},
	legend: { position: 'r', overlay: false, entries: [] },
	plotVisibleOnly: true,
	displayBlanksAs: 'gap',
};

describe('writeChartSpace', () => {
	it('writes a model built from scratch in schema order, and reads it back', () => {
		const xml = writeChartSpace(column, { declarationBreak: '\n' });
		expect(xml).toBe(
			`${DECLARATION}\n${ROOT}<c:roundedCorners val="0"/><c:chart>` +
				'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>Sales &amp; cost</a:t></a:r></a:p></c:rich></c:tx></c:title>' +
				'<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
				'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>' +
				`<c:cat><c:strRef><c:f>'Q&amp;A'!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
				'<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1.50</c:v></c:pt></c:numLit></c:val></c:ser>' +
				'<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>' +
				'<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
				'<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="0.0000001"/></c:scaling><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/><c:crossBetween val="between"/></c:valAx>' +
				'</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>',
		);
		const { chartSpace, issues } = parseChartSpace(xml);
		expect(issues).toEqual([]);
		const { title, ...rest } = withoutSources(chartSpace) as ChartSpace;
		expect(rest).toEqual(withoutSources({ ...column, title: undefined }));
		expect(title?.text).toBe('Sales & cost');
	});

	it('patches only the changed fill, line or effects into a kept spPr', () => {
		const source = `${DECLARATION}${ROOT}<c:chart><c:plotArea><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:round/></a:ln><a:effectLst/><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d></c:spPr></c:plotArea></c:chart></c:chartSpace>`;
		const space = parseChartSpace(source).chartSpace;
		const shape = space.plotArea.spPr!;
		shape.fill = { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } };
		shape.line = { ...shape.line, widthEmu: 19050 };
		expect(writeChartSpace(space, { declarationBreak: '' })).toBe(
			`${DECLARATION}${ROOT}<c:chart><c:plotArea><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="19050" cap="flat" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:round/></a:ln><a:effectLst/><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d></c:spPr></c:plotArea></c:chart></c:chartSpace>`,
		);
		delete shape.line;
		shape.effectsXml = `<a:effectLst xmlns:a="${NS.a}"><a:glow rad="1"/></a:effectLst>`;
		expect(writeChartSpace(space)).toContain(
			'<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst><a:glow rad="1"/></a:effectLst><a:scene3d>',
		);
	});

	it('writes edited text bodies and layouts from the model, kept ones verbatim', () => {
		const source = `${DECLARATION}${ROOT}<c:chart><c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:p><a:r><a:rPr lang="en-US"/><a:t>Kept</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:xMode val="edge"/><c:x val="9.1E-2"/></c:manualLayout></c:layout></c:title><c:plotArea/></c:chart></c:chartSpace>`;
		const space = parseChartSpace(source).chartSpace;
		expect(writeChartSpace(space, { declarationBreak: '' })).toBe(source);
		space.title!.tx!.rich!.paragraphs[0]!.runs[0]!.text = 'Edited';
		space.title!.layout!.x = 0.25;
		expect(writeChartSpace(space)).toContain(
			'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>Edited</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:xMode val="edge"/><c:x val="0.25"/></c:manualLayout></c:layout></c:title>',
		);
	});

	it('writes the style choice, extra root declarations and relationship children', () => {
		const xml = writeChartSpace({
			style: 2,
			c14Style: 102,
			plotArea: { groups: [], axes: [] },
			externalDataRelId: 'rId3',
			externalDataAutoUpdate: false,
			userShapesRelId: 'rId4',
			namespaceDeclarations: [
				{ prefix: 'c16r2', uri: 'http://schemas.microsoft.com/office/drawing/2015/06/chart' },
			],
		});
		expect(xml).toBe(
			`${DECLARATION}\r\n<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">` +
				`<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="c14" xmlns:c14="${NS.c14}"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>` +
				'<c:chart><c:plotArea/></c:chart><c:externalData r:id="rId3"><c:autoUpdate val="0"/></c:externalData><c:userShapes r:id="rId4"/></c:chartSpace>',
		);
		expect(parseChartSpace(xml).chartSpace).toMatchObject({
			style: 2,
			c14Style: 102,
			externalDataAutoUpdate: false,
		});
	});

	it('derives the group element from its kind when the element name is not a chart group', () => {
		const xml = writeChartSpace({
			plotArea: {
				groups: [{ kind: 'pie', element: 'x"/><evil', is3D: true, series: [], axisIds: [] }],
				axes: [],
			},
		});
		expect(xml).toContain('<c:pie3DChart/>');
	});
});

describe('chartNumber', () => {
	it('writes the shortest round-trip decimal without an exponent', () => {
		expect([0, -0, 150, -27, 0.25, 1e-7, -2.5e-8, 1.5e21, 123456789012].map(chartNumber)).toEqual([
			'0',
			'0',
			'150',
			'-27',
			'0.25',
			'0.0000001',
			'-0.000000025',
			'1500000000000000000000',
			'123456789012',
		]);
		expect(() => chartNumber(Number.NaN)).toThrow();
	});
});

describe('stripDeclarations', () => {
	it('drops repeated root declarations unless the fragment rebinds the prefix', () => {
		const bindings = new Map([['a', NS.a]]);
		expect(stripDeclarations(`<a:ln xmlns:a="${NS.a}"><a:x/></a:ln>`, bindings)).toBe(
			'<a:ln><a:x/></a:ln>',
		);
		const rebound = `<a:x xmlns:a="urn:other"><a:y xmlns:a="${NS.a}"/></a:x>`;
		expect(stripDeclarations(rebound, bindings)).toBe(rebound);
	});
});

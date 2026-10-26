import { expect, it } from 'vitest';
import { parseChartSpace } from '../../../chart/index';
import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import { PptxXmlLookupService } from '../services/PptxXmlLookupService';
import type { XmlObject } from '../types';
import { chartFormatFromNeutral } from './chart-from-neutral-format';
import { neutralChartShapeMap } from './chart-from-neutral-format-map';
import { readChartElementFormatsFromXml } from './chart-element-format-read-xml';
import type { ChartTreeReader } from './chart-series-tree-fields';

const reader: ChartTreeReader = {
	xml: new PptxXmlLookupService(),
	parseColor: (node) =>
		(node?.['a:schemeClr'] as XmlObject | undefined)?.['@_val'] === 'accent1'
			? '#204060'
			: undefined,
	resolveTypeface: (value) => value,
	localName: (key) => key.slice(key.lastIndexOf(':') + 1),
};

it('projects neutral effects, alpha, line details and theme colors without changing authored source', () => {
	const shape = {
		fill: {
			kind: 'solid' as const,
			color: {
				kind: 'scheme' as const,
				value: 'accent1',
				transforms: [{ name: 'alpha', value: '50000' }],
			},
		},
		line: {
			widthEmu: 25400,
			cap: 'rnd',
			compound: 'dbl',
			join: { kind: 'round' as const },
			tailEnd: { type: 'triangle' },
		},
		effectsXml:
			'<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="25400" dist="12700" dir="5400000"><a:schemeClr val="accent1"><a:alpha val="25000"/></a:schemeClr></a:outerShdw><a:glow rad="38100"><a:schemeClr val="accent1"/></a:glow><a:softEdge rad="12700"/></a:effectLst>',
	};
	const before = structuredClone(shape);
	expect(chartFormatFromNeutral(shape, reader.parseColor)).toMatchObject({
		fill: { kind: 'solid', color: '#204060', transparency: 50 },
		line: {
			kind: 'automatic',
			width: 2,
			cap: 'rnd',
			compound: 'dbl',
			join: 'round',
			endArrow: 'triangle',
		},
		effects: {
			shadow: { color: '#204060', transparency: 75, blur: 2, distance: 1, angle: 90 },
			glow: { color: '#204060', radius: 3 },
			softEdge: 1,
		},
	});
	expect(shape).toEqual(before);
});

it('maps chart, axis and series primitives to neutral fields while keeping the source tree for opaque details', () => {
	const xml =
		'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:ln w="12700"/></c:spPr></c:ser></c:barChart><c:valAx><c:axId val="2"/><c:scaling/><c:spPr><a:ln w="12700"/></c:spPr></c:valAx></c:plotArea></c:chart></c:chartSpace>';
	const space = parseChartSpace(xml).chartSpace;
	space.plotArea.groups[0].series[0].spPr!.line!.widthEmu = 38100;
	space.plotArea.axes[0].spPr!.line!.widthEmu = 50800;
	const tree = new PptxRuntimeDependencyFactory().createParser().parse(xml) as XmlObject;
	const root = reader.xml.getChildByLocalName(tree, 'chartSpace')!;
	const plot = reader.xml.getChildByLocalName(
		reader.xml.getChildByLocalName(root, 'chart'),
		'plotArea',
	);
	const series = reader.xml.getChildByLocalName(
		reader.xml.getChildByLocalName(plot, 'barChart'),
		'ser',
	)!;
	const formats = readChartElementFormatsFromXml(
		{ chartSpace: root, seriesNodes: [series] },
		{ axes: [{ axisType: 'valAx', axisId: 2 }] },
		reader.localName,
		{
			resolve: (node) => reader.parseColor(node),
			neutralShape: neutralChartShapeMap(space, root, reader),
		},
	);
	expect(formats?.['series:0']?.line?.width).toBe(3);
	expect(formats?.['axis:0']?.line?.width).toBe(4);
});

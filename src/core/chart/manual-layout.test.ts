import { expect, it } from 'vitest';
import { readChartManualLayoutValues, resolveManualLayoutRect } from './manual-layout';
import { writeChartLayoutFormatting } from './write-layout-formatting';
import { readChartFormatting } from './read-formatting';
import { NS, first, parseXml, buildXml } from '../xml';

it('reads the same typed layout values through object-tree and DOM adapters', () => {
	const values: Record<string, unknown> = {
		xMode: 'edge',
		yMode: 'factor',
		x: '0.25',
		y: '-0.1',
		w: '0.5',
		hMode: 'edge',
		layoutTarget: 'inner',
	};
	expect(readChartManualLayoutValues((name) => values[name])).toEqual({
		xMode: 'edge',
		yMode: 'factor',
		x: 0.25,
		y: -0.1,
		width: 0.5,
		heightMode: 'edge',
		layoutTarget: 'inner',
	});
	expect(
		readChartManualLayoutValues(
			(name) =>
				(({ x: ' ', y: 'Infinity', w: true, h: 'no', xMode: 'bogus' }) as Record<string, unknown>)[
					name
				],
		),
	).toBeUndefined();
});

it('resolves edge and factor positions and dimensions using one shared geometry helper', () => {
	const frame = { width: 400, height: 200 },
		auto = { x: 40, y: 20, width: 300, height: 150 };
	expect(
		resolveManualLayoutRect(
			{
				xMode: 'edge',
				x: 0.1,
				yMode: 'edge',
				y: 0.2,
				widthMode: 'edge',
				width: 0.9,
				heightMode: 'edge',
				height: 0.7,
			},
			frame,
			auto,
		),
	).toEqual({ x: 40, y: 40, width: 320, height: 100 });
	expect(
		resolveManualLayoutRect({ x: 0.1, y: -0.05, width: 0.5, height: 0.25 }, frame, auto),
	).toEqual({ x: 80, y: 10, width: 200, height: 50 });
	expect(resolveManualLayoutRect({}, frame, auto)).toBeUndefined();
});

it('preserves layout extensions and overlay for every chart region in schema order', () => {
	const source = parseXml(
		`<c:chartSpace xmlns:c="${NS.c}"><c:chart><c:title><c:tx/><c:layout><c:manualLayout><c:x val="0.2"/><c:extLst><c:ext uri="unknown"/></c:extLst></c:manualLayout></c:layout><c:overlay val="1"/></c:title><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:w val="0.5"/></c:manualLayout></c:layout></c:plotArea><c:legend><c:legendPos val="r"/><c:layout><c:manualLayout><c:y val="0.25"/></c:manualLayout></c:layout><c:overlay val="1"/></c:legend></c:chart></c:chartSpace>`,
	);
	const target = parseXml(
		`<c:chartSpace xmlns:c="${NS.c}"><c:chart><c:title><c:tx/><c:spPr/></c:title><c:plotArea><c:layout/></c:plotArea><c:legend><c:legendPos val="r"/><c:spPr/></c:legend></c:chart></c:chartSpace>`,
	);
	expect(
		writeChartLayoutFormatting(
			target.documentElement,
			readChartFormatting(source.documentElement)!,
		),
	).toBe(true);
	const chart = first(target.documentElement, 'chart', NS.c);
	for (const region of ['title', 'plotArea', 'legend'])
		expect(buildXml(first(first(chart, region, NS.c), 'layout', NS.c)!)).toBe(
			buildXml(
				first(first(first(source.documentElement, 'chart', NS.c), region, NS.c), 'layout', NS.c)!,
			),
		);
	expect(buildXml(first(chart, 'title', NS.c)!)).toMatch(
		/<c:tx\/><c:layout>.*<c:overlay val="1"\/><c:spPr\/>/,
	);
});

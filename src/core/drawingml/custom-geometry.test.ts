import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { parseDrawingGeometry } from './drawing-geometry';
import { parseCustomGeometry } from './custom-geometry';
import { drawingPathData } from './path-data';
import { parseDiagramDrawing } from '../diagram/drawing';
import { custom } from './custom-geometry.test-support';
const ns = 'http://schemas.openxmlformats.org/drawingml/2006/main';
describe('shared guide-driven custom geometry', () => {
	it('evaluates guides at the current frame and retains ordered commands and per-path space', () => {
		const element = parseXml(custom).documentElement;
		const paths = parseCustomGeometry(element, 200, 80)!;
		expect(paths[0]?.commands[0]).toEqual({ op: 'M', x: 50, y: 0 });
		expect(paths[1]).toMatchObject({ width: 100, height: 100, fill: 'none', stroke: false });
		expect(drawingPathData(paths[1]!, 200, 80)).toContain('A 100 40');
		expect(parseCustomGeometry(element, 400, 100)?.[0]?.commands[0]).toEqual({
			op: 'M',
			x: 100,
			y: 0,
		});
	});
	it('uses the DrawingML reader and reports unresolved guides or unknown commands', () => {
		const read = (source: string) =>
			parseDrawingGeometry(
				parseXml(
					`<a:spPr xmlns:a="${ns}"><a:xfrm><a:off x="0" y="0"/><a:ext cx="200" cy="80"/></a:xfrm>${source}</a:spPr>`,
				).documentElement,
			);
		expect(read(custom)).toMatchObject({
			geometry: 'custom',
			paths: [
				{
					commands: [
						{ op: 'M', x: 50, y: 0 },
						{ op: 'L', x: 200, y: 80 },
						{ op: 'L', x: 0, y: 80 },
						{ op: 'Z' },
					],
				},
				{},
			],
		});
		expect(read(custom.replace('x="x"', 'x="missing"')).issue).toBe('CUSTOM_GEOMETRY_UNSUPPORTED');
		expect(read(custom.replace('<a:close/>', '<a:future/>')).issue).toBe(
			'CUSTOM_GEOMETRY_UNSUPPORTED',
		);
	});
});

it('evaluates cached SmartArt paths through the same shared reader', () => {
	const drawing = parseDiagramDrawing(
		`<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="${ns}"><dsp:spTree><dsp:sp modelId="node"><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="200" cy="80"/></a:xfrm>${custom}</dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`,
	);
	expect(drawing.issues).toEqual([]);
	expect(drawing.shapes[0]?.paths?.[0]?.commands[0]).toEqual({ op: 'M', x: 50, y: 0 });
});

it('keeps shape guide coordinates independent of differing local path dimensions', () => {
	const source = custom.replace('<a:path>', '<a:path w="100" h="40">');
	const paths = parseCustomGeometry(parseXml(source).documentElement, 200, 80)!;
	expect(paths[0]).toMatchObject({
		width: 100,
		height: 40,
		commands: [
			{ op: 'M', x: 50, y: 0 },
			{ op: 'L', x: 200, y: 80 },
			{ op: 'L', x: 0, y: 80 },
			{ op: 'Z' },
		],
	});
	expect(paths[1]).toMatchObject({ width: 100, height: 100 });
	expect(drawingPathData(paths[0]!, 200, 80)).toContain('M100 0 L400 160');
});

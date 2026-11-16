import { describe, expect, it } from 'vitest';
import type { DiagramDrawingShape } from 'ooxml-core/diagram';
import { renderDiagramDrawing } from './smartart-svg';

const E = 9525;
const shape = (over: Partial<DiagramDrawingShape> = {}): DiagramDrawingShape => ({
	modelId: 'node',
	frame: { x: 20 * E, y: 20 * E, width: 160 * E, height: 80 * E },
	geometry: 'rect',
	has3d: false,
	fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4', transforms: [] } },
	...over,
});
const bevel = '<a:sp3d prstMaterial="plastic"><a:bevelT w="63500" h="25400"/></a:sp3d>';

describe('approximate SmartArt 3D', () => {
	it("lights a bevel from the label's light rig instead of drawing it flat", () => {
		const { svg, report } = renderDiagramDrawing(document, {
			issues: [],
			shapes: [shape({ has3d: true, shape3d: { shape3dXml: bevel } })],
		});
		expect(report).toMatchObject({ flattened3d: 0, approximated3d: 1 });
		const wrap = svg.querySelector('[data-bevel]')!;
		expect(wrap.getAttribute('data-bevel')).toBe('circle');
		const filter = svg.querySelector(wrap.getAttribute('filter')!.slice(4, -1))!;
		expect(filter.querySelector('feSpecularLighting feDistantLight')!.getAttribute('azimuth')).toBe(
			'270',
		);
		expect(svg.querySelector('[data-flattened-3d]')).toBeNull();
	});

	it("projects a Scene style's spatial camera and shows its extrusion", () => {
		const { svg, report } = renderDiagramDrawing(document, {
			issues: [],
			extentEmu: { width: 400 * E, height: 200 * E },
			scene3d: { cameraPreset: 'isometricOffAxis2Left' },
			shapes: [
				shape({
					has3d: true,
					view3d: { shape: { extrusionHeight: 25400 * 4, presetMaterial: 'matte' } },
				}),
				shape({
					modelId: 'flat',
					frame: { x: 220 * E, y: 20 * E, width: 160 * E, height: 80 * E },
				}),
			],
		});
		expect(report.camera).toBe('isometricOffAxis2Left');
		expect(report.approximated3d).toBe(2);
		const groups = [...svg.querySelectorAll<SVGGElement>('g[data-model-id]')];
		for (const group of groups) expect(group.getAttribute('transform')).toMatch(/^matrix\(/u);
		// A parallel camera turned to the left shears the frame: its x edge rises or falls.
		const [, b] = /^matrix\(([-\d.]+) ([-\d.]+)/u
			.exec(groups[0]!.getAttribute('transform')!)!
			.slice(1)
			.map(Number);
		expect(Math.abs(b!)).toBeGreaterThan(0.05);
		const side = groups[0]!.querySelector('[data-extrusion]')!;
		expect(side.children.length).toBeGreaterThan(1);
		expect(groups[1]!.querySelector('[data-extrusion]')).toBeNull();
	});

	it('hides an extrusion seen straight on and keeps the plain placement', () => {
		const { svg, report } = renderDiagramDrawing(document, {
			issues: [],
			shapes: [shape({ has3d: true, view3d: { shape: { extrusionHeight: 76200 } } })],
		});
		expect(svg.querySelector('[data-extrusion]')).toBeNull();
		expect(svg.querySelector('g[data-model-id]')!.getAttribute('transform')).toMatch(
			/^translate\(/u,
		);
		expect(report.flattened3d).toBe(1);
	});
});

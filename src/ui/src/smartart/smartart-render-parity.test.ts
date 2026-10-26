import { expect, it } from 'vitest';
import { evaluatePresetShape, ST_SHAPE_TYPE_VALUES } from 'ooxml-core/geometry';
import type { DiagramDrawingShape } from 'ooxml-core/diagram';
import { drawingPathData } from './smartart-geometry';
import { renderDiagramDrawing } from './smartart-svg';

const E = 9525;
const shape = (over: Partial<DiagramDrawingShape> = {}): DiagramDrawingShape => ({
	modelId: 'node',
	frame: { x: 0, y: 0, width: 160 * E, height: 80 * E },
	geometry: 'rect',
	has3d: false,
	...over,
});
const color = (value: string) => ({ kind: 'srgb' as const, value, transforms: [] });
const render = (shapes: DiagramDrawingShape[]) =>
	renderDiagramDrawing(document, { shapes, issues: [] });

it('shares preset geometry and authored adjustments with the core evaluator across the corpus', () => {
	let evaluated = 0;
	for (const name of ST_SHAPE_TYPE_VALUES) {
		const expected = evaluatePresetShape(name, 160, 80, { adj: 40000 });
		const result = render([shape({ geometry: name, adjustments: { adj: 40000 } })]);
		if (!expected) {
			expect(result.report.approximatedGeometries).toContain(name);
			continue;
		}
		evaluated++;
		expect(result.report.approximatedGeometries, name).toEqual([]);
		if (!['rect', 'ellipse'].includes(name))
			expect(
				[...result.svg.querySelectorAll('path')].map((path) => path.getAttribute('d')),
				name,
			).toEqual(expected.paths.map((path) => path.d));
	}
	expect(evaluated).toBeGreaterThan(170);
});

it('renders custom arcs including full turns, reversed sweeps and closed-path pen restoration', () => {
	const path = drawingPathData(
		{
			width: 100,
			height: 100,
			commands: [
				{ op: 'M', x: 100, y: 50 },
				{ op: 'A', wR: 50, hR: 50, startAngle: 0, swingAngle: 360 },
				{ op: 'Z' },
				{ op: 'A', wR: 50, hR: 50, startAngle: 0, swingAngle: -90 },
			],
		},
		200,
		100,
	);
	expect(path.match(/A /g)).toHaveLength(3);
	expect(path).not.toMatch(/NaN|Infinity/);
	expect(path).toContain('100 50');
});

it('paints alpha gradients with unique local servers and retains open-path paint flags', () => {
	const fill = {
		kind: 'gradient' as const,
		angle: 90,
		stops: [
			{ position: 0, color: color('FF0000') },
			{
				position: 100,
				color: { ...color('0000FF'), transforms: [{ name: 'alpha', value: '30000' }] },
			},
		],
	};
	const first = render([shape({ geometry: 'arc', fill })]);
	const second = render([shape({ fill })]);
	expect(first.svg.querySelector('linearGradient')!.id).not.toBe(
		second.svg.querySelector('linearGradient')!.id,
	);
	expect(first.svg.querySelector('stop[offset="100%"]')!.getAttribute('stop-color')).toBe(
		'rgba(0, 0, 255, 0.3)',
	);
	expect(first.svg.querySelector('path[fill="none"]')).not.toBeNull();
	expect(first.report.approximatedFills).toEqual([]);
});

it('preserves paragraph breaks and distinct run styles when wrapping text', () => {
	const result = render([
		shape({
			text: {
				anchor: 't',
				text: 'Bold plain\nSecond',
				paragraphs: [
					{
						align: 'l',
						runs: [
							{ text: 'Bold ', bold: true, sizePt: 16 },
							{ text: 'plain', italic: true, sizePt: 16, color: color('FF0000') },
						],
					},
					{ align: 'r', runs: [{ text: 'Second', underline: true, sizePt: 10 }] },
				],
			},
		}),
	]);
	const text = result.svg.querySelector('text')!;
	expect(text.querySelector('tspan[font-weight="700"]')!.textContent).toBe('Bold ');
	expect(text.querySelector('tspan[font-style="italic"]')!.getAttribute('fill')).toBe('#FF0000');
	expect(text.querySelector('tspan[text-decoration="underline"]')!.textContent).toBe('Second');
	const lines = [...text.children];
	expect(lines).toHaveLength(2);
	expect(lines[0]!.getAttribute('y')).not.toBe(lines[1]!.getAttribute('y'));
	expect(lines[0]!.getAttribute('text-anchor')).toBe('start');
	expect(lines[1]!.getAttribute('text-anchor')).toBe('end');
});

it('reports unsupported fills, geometry and optional 3D honestly', () => {
	const result = render([
		shape({ geometry: 'foreign-shape', has3d: true, fill: { kind: 'picture' } }),
	]);
	expect(result.report).toMatchObject({
		approximatedGeometries: ['foreign-shape'],
		approximatedFills: ['picture'],
		flattened3d: 1,
	});
});

it('uses the shared authored shadow projection and reports effects it cannot paint', () => {
	const result = render([
		shape({
			effects: {
				effects: [
					{
						kind: 'outerShadow',
						color: { ...color('000000'), transforms: [{ name: 'alpha', value: '25000' }] },
						blurRadiusEmu: 4 * E,
						distanceEmu: 6 * E,
						directionAngle: 90 * 60000,
					},
					{ kind: 'unsupported', sourceXml: '<a:foreignFx/>' },
				],
			},
		}),
	]);
	const shadow = result.svg.querySelector('feDropShadow')!;
	expect(Number(shadow.getAttribute('dy'))).toBeCloseTo(6);
	expect(Number(shadow.getAttribute('dx'))).toBeCloseTo(0);
	expect(shadow.getAttribute('stdDeviation')).toBe('2');
	expect(shadow.getAttribute('flood-opacity')).toBe('0.25');
	expect(result.report.unappliedEffects).toEqual(['unsupported']);
});

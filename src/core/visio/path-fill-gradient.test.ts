import { expect, it } from 'vitest';
import { pathFillGradient } from './path-fill-gradient';
import { demoDocument } from './ui/demo-document';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import { normalizeVisioPageGeometry } from './page-scale';
import type { VisioGeometry } from './model';

const stops = [
	{ offset: 0, color: '#ff0000', opacity: 1 },
	{ offset: 1, color: '#0000ff', opacity: 1 },
];
const geometry = (path: string): VisioGeometry => ({ path, fill: true, stroke: true });
const rectangle = 'M 0 0 L 2 0 L 2 1 L 0 1 L 0 0';
const paint = (path: string) => pathFillGradient([geometry(path)], 2, 1, stops);

it('uses the bounding-box center and perpendicular outline endpoints for a triangle', () => {
	const gradient = paint('M 0 0 L 2 0 L 1 1 L 0 0');
	expect(gradient).toMatchObject({ type: 'regions', coordinateSpace: 'shape' });
	if (gradient?.type !== 'regions') throw new Error('Missing polygon gradient');
	expect(gradient.regions).toHaveLength(3);
	expect(gradient.regions.map((region) => region.end)).toEqual([
		[0.5, 0],
		[0.7, 0.6],
		[0.3, 0.6],
	]);
	expect(
		gradient.regions.every((region) => region.start[0] === 0.5 && region.start[1] === 0.5),
	).toBe(true);
	expect(paint('M 0 0 L 1 1 L 2 0 L 0 0')?.type).toBe('regions');
});

it('does not create negative or degenerate faces at a reentrant notch', () => {
	const gradient = paint('M 0 0 L 2 0 L 2 0.4 L 1 0.4 L 1 1 L 0 1 L 0 0');
	if (gradient?.type !== 'regions') throw new Error('Missing notched gradient');
	expect(gradient.regions).toHaveLength(4);
	expect(gradient.regions.map((region) => region.end)).toEqual([
		[0.5, 0],
		[1, 0.5],
		[0.5, 1],
		[0, 0.5],
	]);
});

it('retains an analytic normalized ellipse and rejects unsupported curved/offset outlines', () => {
	expect(paint('M 2 0.5 A 1 0.5 0 0 1 0 0.5 A 1 0.5 0 0 1 2 0.5 Z')).toEqual({
		type: 'radial',
		center: [0.5, 0.5],
		radius: 0.5,
		stops,
	});
	for (const path of [
		'M 0 0 C 2 0 2 1 0 1 Z',
		'M 1.5 0.5 A 0.5 0.5 0 0 1 0.5 0.5 A 0.5 0.5 0 0 1 1.5 0.5 Z',
		'M 0.1 0 L 2 0 L 2 1 L 0.1 1 Z',
		'M 0 0 L 2 0 L 2 1',
		'M 0 0 L 2 1 L 0 1 L 2 0 Z',
		'M 0 0 L 2 0 L 1 0 L 2 1 L 0 1 Z',
		'M 0 0 L 2 0 L 2 1 L 0 1 Z M 0.5 0.2 L 1 0.2 L 1 0.5 Z',
	])
		expect(paint(path), path).toBeUndefined();
	expect(pathFillGradient([geometry(rectangle), geometry(rectangle)], 2, 1, stops)).toBeUndefined();
});

it('preserves normalized fan coordinates in snapshots and page scaling while counting each paint', () => {
	const model = structuredClone(demoDocument),
		shape = model.pages[0]!.shapes[0]!;
	shape.style.fillGradient = paint('M 0 0 L 2 0 L 1 1 L 0 0')!;
	model.pages = [model.pages[0]!];
	model.pages[0]!.shapes = [shape];
	const copy = copySnapshotScene(model);
	normalizeVisioPageGeometry(copy.pages[0]!.shapes, 0.5, () => {});
	expect(copy.pages[0]!.shapes[0]!.style.fillGradient).toEqual(shape.style.fillGradient);
	assertViewableDocument(copy);
	const gradient = shape.style.fillGradient;
	if (gradient.type !== 'regions') throw new Error('Missing gradient');
	gradient.interpolation = 'sigma-gamma22';
	model.pages[0]!.shapes = Array.from({ length: 131 }, (_, id) => ({ ...shape, id: String(id) }));
	expect(() => assertViewableDocument(model)).toThrow('gradient stop limits');
	model.pages[0]!.shapes = [shape];
	gradient.regions[0]!.angle = 1;
	expect(() => assertViewableDocument(model)).toThrow('gradient region angle');
});

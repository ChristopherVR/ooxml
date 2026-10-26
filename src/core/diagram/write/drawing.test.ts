import { expect, it } from 'vitest';
import { writeGeneratedDiagramDrawing } from './drawing';
import { parseDiagramDrawing } from '../drawing';
import { previewDiagramLayout } from '../layout-preview';
import { builtinLayoutIdForPreset, loadBuiltinSmartArtLayoutXml } from '../builtin';
import type { DiagramDrawing } from '../types';

it('serializes an interpreted 2D layout independently of Office cache regeneration', () => {
	const preview = previewDiagramLayout({
		extentEmu: { width: 600 * 9525, height: 350 * 9525 },
		nodes: [
			{ id: 'a', text: 'Plan' },
			{ id: 'b', text: 'Build' },
		],
		source: {
			layoutXml: loadBuiltinSmartArtLayoutXml(builtinLayoutIdForPreset('basicChevronProcess')!),
		},
	});
	const result = writeGeneratedDiagramDrawing(preview.drawing!);
	const drawing = parseDiagramDrawing(result.xml);
	expect(drawing.shapes.map((shape) => [shape.frame, shape.geometry, shape.text?.text])).toEqual(
		preview.drawing!.shapes.map((shape) => [shape.frame, shape.geometry, shape.text?.text]),
	);
	expect(result.modelIds.size).toBe(drawing.shapes.length);
});

it('preserves custom arcs, alpha, rich runs and typed effects in a generated cache', () => {
	const drawing: DiagramDrawing = {
		issues: [],
		shapes: [
			{
				modelId: 'a',
				geometry: 'custom',
				has3d: false,
				frame: { x: 0, y: 0, width: 952500, height: 476250 },
				paths: [
					{
						width: 100,
						height: 100,
						commands: [
							{ op: 'M', x: 100, y: 50 },
							{ op: 'A', wR: 50, hR: 50, startAngle: 0, swingAngle: 360 },
						],
					},
				],
				fill: {
					kind: 'solid',
					color: {
						kind: 'scheme',
						value: 'accent2',
						transforms: [{ name: 'alpha', value: '50000' }],
					},
				},
				body: {
					paragraphs: [
						{
							items: [
								{ kind: 'run', text: 'Bold', properties: { bold: true } },
								{ kind: 'run', text: ' italic', properties: { italic: true } },
							],
						},
					],
				},
				effects: {
					effects: [
						{
							kind: 'outerShadow',
							distanceEmu: 9525,
							color: { kind: 'srgb', value: '000000', transforms: [] },
						},
					],
				},
			},
		],
	};
	const saved = parseDiagramDrawing(writeGeneratedDiagramDrawing(drawing).xml);
	expect(saved.shapes[0]?.paths).toEqual(drawing.shapes[0]?.paths);
	expect(saved.shapes[0]?.fill).toEqual(drawing.shapes[0]?.fill);
	expect(saved.shapes[0]?.body?.paragraphs[0]?.items).toEqual(
		drawing.shapes[0]?.body?.paragraphs[0]?.items,
	);
	expect(saved.shapes[0]?.effects?.effects[0]).toMatchObject({
		kind: 'outerShadow',
		distanceEmu: 9525,
	});
	expect(() =>
		writeGeneratedDiagramDrawing({ ...drawing, shapes: [{ ...drawing.shapes[0]!, has3d: true }] }),
	).toThrow('3D');
});

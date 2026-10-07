import { describe, expect, it } from 'vitest';
import { visioPageEditToDrawing } from './page-edit.js';
import { demoDocument } from './demo-document.js';
import { copySnapshotScene } from './snapshot-scene.js';
import { assertViewableDocument } from './scene-validation.js';

describe('page/drawing edit coordinate boundary', () => {
	it('converts creation, move and resize without altering text/delete edits or caller state', () => {
		const page = { ...demoDocument.pages[0]!, drawingToPageScale: 0.5 };
		const create = {
			type: 'create-rectangle' as const,
			pageId: page.id,
			shapeId: '99',
			x: 1,
			y: 2,
			width: 3,
			height: 4,
			text: 'kept',
		};
		expect(visioPageEditToDrawing(page, create)).toEqual({
			...create,
			x: 2,
			y: 4,
			width: 6,
			height: 8,
		});
		expect(create.x).toBe(1);
		expect(
			visioPageEditToDrawing(page, {
				type: 'move-shape',
				pageId: page.id,
				shapeId: '1',
				x: 3,
				y: 4,
			}),
		).toMatchObject({ x: 6, y: 8 });
		expect(
			visioPageEditToDrawing(page, {
				type: 'resize-shape',
				pageId: page.id,
				shapeId: '1',
				width: 3,
				height: 4,
			}),
		).toMatchObject({ width: 6, height: 8 });
		const text = {
			type: 'replace-plain-text' as const,
			pageId: page.id,
			shapeId: '1',
			text: 'same',
		};
		expect(visioPageEditToDrawing(page, text)).toBe(text);
	});
	it('preserves the page scale through bounded scene snapshots', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.drawingToPageScale = 0.5;
		assertViewableDocument(model);
		expect(copySnapshotScene(model).pages[0]!.drawingToPageScale).toBe(0.5);
	});
	it.each([0, -1, NaN, Infinity])(
		'rejects invalid scale %s at display and edit boundaries',
		(scale) => {
			const model = structuredClone(demoDocument);
			model.pages[0]!.drawingToPageScale = scale;
			expect(() => assertViewableDocument(model)).toThrow(/drawing scale/);
			expect(() =>
				visioPageEditToDrawing(model.pages[0]!, {
					type: 'delete-shape',
					pageId: model.pages[0]!.id,
					shapeId: '1',
				}),
			).toThrow();
		},
	);
});

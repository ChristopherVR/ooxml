/**
 * Regression: a structural SmartArt edit (Enter / Backspace in the text pane) clears the drawing
 * shapes, and the rebuild every binding runs must keep the picture fills of a picture layout
 * instead of falling back to the algorithmic engine, which drops every image.
 */
import { readFileSync } from 'node:fs';

import { PptxHandler, addSmartArtNode } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { withNativePictureLayout } from '../../../../core/diagram/test-support/picture-layout';
import { rebuildDrawingShapesIfCleared } from './smartart-reflow-to-shapes';

describe('rebuildDrawingShapesIfCleared - picture layouts', () => {
	it.each(['picture-accent-list--flat3', 'picture-grid--flat3'])(
		'keeps the image fills of %s after a node is added',
		async (layout) => {
			const original = readFileSync(
				new URL(
					`../../../../core/pptx/__tests__/fixtures/smartart-gallery/${layout}.pptx`,
					import.meta.url,
				),
			);
			const bytes = await withNativePictureLayout(original, 'ppt', layout);
			const { slides } = await new PptxHandler().load(bytes);
			const element = slides[0]!.elements.find((item) => item.type === 'smartArt');
			if (element?.type !== 'smartArt' || !element.smartArtData) {
				throw new Error('Missing SmartArt');
			}
			const edited = addSmartArtNode(
				element.smartArtData,
				'Fourth',
				element.smartArtData.nodes[0]!.id,
			);
			expect(edited.drawingShapes ?? []).toHaveLength(0);
			const rebuilt = rebuildDrawingShapesIfCleared(
				{ ...edited, drawingShapes: [] },
				edited.layout,
				[],
				'flat',
				element.id,
				{ width: element.width, height: element.height },
			);
			expect(rebuilt.drawingShapes?.filter((shape) => shape.fillImageUrl)).toHaveLength(3);
			const captions = (rebuilt.drawingShapes ?? []).map((shape) => shape.text).filter(Boolean);
			expect(captions).toEqual(expect.arrayContaining(['Alpha', 'Gamma', 'Fourth']));
		},
	);
});

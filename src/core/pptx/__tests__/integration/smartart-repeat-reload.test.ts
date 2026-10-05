import { describe, expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxSlide, SmartArtPptxElement } from '../../core/types';
import { readCorpusFixture } from './real-world-corpus-helpers';

function diagrams(slides: PptxSlide[]) {
	return slides
		.flatMap((slide) => slide.elements)
		.filter((element): element is SmartArtPptxElement => element.type === 'smartArt')
		.map((element) => ({
			layout: element.smartArtData?.resolvedLayoutType,
			nodes: element.smartArtData?.nodes.map((node) => ({
				id: node.id,
				text: node.text,
				parentId: node.parentId,
			})),
			shapes: element.smartArtData?.drawingShapes?.map((shape) => ({
				shapeType: shape.shapeType,
				x: shape.x,
				y: shape.y,
				width: shape.width,
				height: shape.height,
				text: shape.text,
			})),
		}));
}

describe('PowerPoint SmartArt repeated save and reload', () => {
	it.each([
		'smartart-chart-table-mix.pptx',
		'smartart-orgchart-assistants.pptx',
		'smartart-orgchart-nested-hang.pptx',
	])('preserves %s across two saves and reloads', async (fixture) => {
		let handler = new PptxHandler();
		let loaded = await handler.load(readCorpusFixture(fixture));
		const expected = diagrams(loaded.slides);
		expect(expected.length).toBeGreaterThan(0);
		for (let round = 0; round < 2; round++) {
			const saved = await handler.save(loaded.slides);
			handler.dispose();
			handler = new PptxHandler();
			loaded = await handler.load(
				saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
			);
			expect(diagrams(loaded.slides)).toStrictEqual(expected);
		}
		handler.dispose();
	});
});

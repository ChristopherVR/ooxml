import { describe, expect, it } from 'vitest';
import { approximateDiagramDrawingFromNodes } from './approximate';

describe('approximateDiagramDrawingFromNodes', () => {
	const extentEmu = { width: 400 * 9525, height: 200 * 9525 };

	it('lays the nodes out inside the frame with colourful accents when the colours ask', () => {
		const drawing = approximateDiagramDrawingFromNodes({
			extentEmu,
			nodes: [
				{ id: 'a', text: 'A' },
				{ id: 'b', text: 'B' },
			],
			family: 'list',
			colorsId: 'urn:microsoft.com/office/officeart/2005/8/colors/colorful1',
		})!;
		expect(drawing.shapes.map((shape) => shape.fill)).toEqual([
			{ kind: 'solid', color: { kind: 'scheme', value: 'accent1', transforms: [] } },
			{ kind: 'solid', color: { kind: 'scheme', value: 'accent2', transforms: [] } },
		]);
	});

	it('draws nothing for an empty frame or no nodes', () => {
		expect(approximateDiagramDrawingFromNodes({ extentEmu, nodes: [] })).toBeUndefined();
		expect(
			approximateDiagramDrawingFromNodes({
				extentEmu: { width: 0, height: 0 },
				nodes: [{ id: 'a', text: 'A' }],
			}),
		).toBeUndefined();
	});
});

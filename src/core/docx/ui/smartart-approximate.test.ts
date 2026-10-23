import { describe, expect, it } from 'vitest';
import { createDocxDiagram } from '../diagram-edit';
import { approximateDiagramDrawing } from './smartart-approximate';

const diagram = (family: 'list' | 'hierarchy', items: string[]) =>
	createDocxDiagram({
		layoutId: `urn:test/${family}`,
		family,
		colorsId: 'urn:microsoft.com/office/officeart/2005/8/colors/accent1_2',
		quickStyleId: 'urn:test/style',
		items,
		widthPx: 400,
		heightPx: 200,
		partSuffix: '-t',
	});

describe('approximateDiagramDrawing', () => {
	it('lays every node out inside the frame, empty ones included, with themed fills', () => {
		const drawing = approximateDiagramDrawing(diagram('list', ['One', '', 'Three']))!;
		expect(drawing.shapes).toHaveLength(3);
		for (const shape of drawing.shapes) {
			expect(shape.frame.x + shape.frame.width).toBeLessThanOrEqual(400 * 9525);
			expect(shape.frame.y + shape.frame.height).toBeLessThanOrEqual(200 * 9525);
			expect(shape.fill).toEqual({
				kind: 'solid',
				color: { kind: 'scheme', value: 'accent1', transforms: [] },
			});
		}
		expect(drawing.shapes.map((shape) => shape.text?.text ?? '')).toEqual(['One', '', 'Three']);
	});

	it('places a hierarchy child below its parent', () => {
		const drawing = approximateDiagramDrawing(diagram('hierarchy', ['Top', 'Left', 'Right']))!;
		const [top, left] = drawing.shapes;
		expect(left!.frame.y).toBeGreaterThan(top!.frame.y);
	});

	it('draws nothing without nodes', () => {
		expect(approximateDiagramDrawing(diagram('list', []))).toBeUndefined();
	});
});

import { expect, it } from 'vitest';
import type { VisioPage, VisioShape } from '../model';
import { visioSelectionShapes } from './selection-connectors';

const shape = (id: string, kind: VisioShape['kind'], children: VisioShape[] = []) =>
	({ id, kind, children }) as unknown as VisioShape;
const link = (fromShapeId: string, fromCell: string, toShapeId: string) => ({
	fromShapeId,
	fromCell,
	toShapeId,
	toCell: 'PinX',
});
const page = (connectors: VisioPage['connectors']) =>
	({
		shapes: [
			shape('1', 'shape'),
			shape('2', 'group', [shape('5', 'shape')]),
			shape('3', 'connector'),
			shape('4', 'shape'),
		],
		connectors,
	}) as unknown as VisioPage;

it('keeps a selection without connectors as it is', () => {
	expect(visioSelectionShapes(page([]), ['1', '2'])).toEqual(['1', '2']);
});

it('drops a connector whose both ends are glued inside the selection', () => {
	const both = page([link('3', 'BeginX', '1'), link('3', 'EndX', '5')]);
	// Shape 5 is a member of the selected group 2.
	expect(visioSelectionShapes(both, ['1', '2', '3'])).toEqual(['1', '2']);
});

it('refuses a connector with a free end, an end outside the selection, or alone', () => {
	expect(visioSelectionShapes(page([link('3', 'BeginX', '1')]), ['1', '2', '3'])).toBeUndefined();
	const outside = page([link('3', 'BeginX', '1'), link('3', 'EndX', '4')]);
	expect(visioSelectionShapes(outside, ['1', '2', '3'])).toBeUndefined();
	expect(visioSelectionShapes(outside, ['3'])).toBeUndefined();
});

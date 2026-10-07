import { expect, it } from 'vitest';
import { analyzeVisioFormula } from './formula';
import { editVsdx } from './edit';
import { cell, fixture, rectangle, section, row, shape } from './test-fixtures';

it('indexes polyline arguments without hiding nested dynamic references', () => {
	expect(analyzeVisioFormula('POLYLINE(0,0,Sheet.1!PinX,0)')).toMatchObject({
		dynamic: false,
		unsupportedFunctions: ['POLYLINE'],
		references: [{ shapeId: '1', cell: 'PinX' }],
	});
	expect(analyzeVisioFormula('POLYLINE(0,0,GETREF(PinX),0)').dynamic).toBe(true);
});

it('still rejects edits that would require recalculating unsupported polyline data', async () => {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape(
					'1',
					rectangle + cell('Width', 2) + cell('Height', 1) + cell('PinX', 2) + cell('PinY', 1),
				)}${shape(
					'2',
					rectangle + section('User', row(0, '', cell('Value', 1, 'POLYLINE(0,0,Sheet.1!PinX,0)'))),
				)}</Shapes>`,
			},
		],
	});
	await expect(
		editVsdx(bytes, [{ type: 'move-shape', pageId: '0', shapeId: '1', x: 3, y: 1 }]),
	).rejects.toThrow('unsupported functions');
});

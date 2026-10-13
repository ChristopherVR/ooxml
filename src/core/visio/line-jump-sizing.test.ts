import { describe, expect, it } from 'vitest';
import evidence from './__fixtures__/native-line-jump-sizing.json';
import { parseVsdx } from './parser';
import { visioLineJumpPaths } from './line-jumps';
import { cell, fixture, shape } from './test-fixtures';
import { copySnapshotScene } from './ui/snapshot-scene';
import { editVsdx } from './edit';

describe('native Visio line jump sizing', () => {
	it.each(evidence.cases)(
		'matches native %s factors %s/%s spacing %s/%s scale %s/%s weight %s',
		async (...row) => {
			const [
				direction,
				factorX,
				factorY,
				spacingX,
				spacingY,
				pageScale,
				drawingScale,
				weight,
				radius,
			] = row;
			const line = (id: string, x1: number, y1: number, x2: number, y2: number) =>
				shape(
					id,
					cell('BeginX', x1) +
						cell('BeginY', y1) +
						cell('EndX', x2) +
						cell('EndY', y2) +
						cell('LineWeight', Number(weight) / 72) +
						cell('ShapeRouteStyle', 16),
				);
			const model = await parseVsdx(
				await fixture({
					pages: [
						{
							id: '0',
							pageCells:
								cell('LineJumpCode', direction === 'h' ? 1 : 2) +
								cell('LineJumpFactorX', Number(factorX)) +
								cell('LineJumpFactorY', Number(factorY)) +
								cell('LineToLineX', Number(spacingX)) +
								cell('LineToLineY', Number(spacingY)) +
								cell('PageScale', Number(pageScale)) +
								cell('DrawingScale', Number(drawingScale)),
							contents: `<Shapes>${line('1', 1, 2, 5, 2)}${line('2', 3, 0, 3, 4)}</Shapes>`,
						},
					],
				}),
			);
			for (const page of [model.pages[0]!, copySnapshotScene(model).pages[0]!]) {
				const paths = [...visioLineJumpPaths(page).values()].flat();
				if (radius === null) expect(paths).toEqual([]);
				else {
					expect(paths).toHaveLength(1);
					const actual = Number(paths[0]!.match(/A ([\d.]+)/)?.[1]) * 72;
					expect(actual).toBeCloseTo(Number(radius), 3);
				}
			}
		},
	);
	it('saves factors and routing distances and refuses invalid values', async () => {
		const bytes = await fixture();
		const edit = {
			type: 'set-page-layout' as const,
			pageId: '0',
			lineJumpFactorX: 0,
			lineJumpFactorY: 2,
			lineToLineX: 0.25,
			lineToLineY: 0.5,
		};
		const saved = await editVsdx(bytes, [edit]);
		expect((await parseVsdx(saved.bytes)).pages[0]!.layout).toMatchObject({
			lineJumpFactorX: 0,
			lineJumpFactorY: 2,
			lineToLineX: 0.25,
			lineToLineY: 0.5,
		});
		for (const patch of [{ lineJumpFactorX: 11 }, { lineJumpFactorY: -1 }, { lineToLineX: -0.1 }])
			await expect(editVsdx(bytes, [{ ...edit, ...patch }])).rejects.toMatchObject({
				code: 'INVALID_EDIT',
			});
	});
});

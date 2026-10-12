import { describe, expect, it } from 'vitest';
import { parseVsdx, visioLineJumpPaths } from 'ooxml-core/visio';
import { copySnapshotScene } from 'ooxml-core/visio/ui';
import { createPrintSnapshot } from './print-snapshot';
import { fixture, shape, cell } from '../../../core/visio/test-fixtures';

describe('Visio line jumps in print snapshots', () => {
	it.each([true, false])(
		'preserves screen jumps in print with explicit route metadata: %s',
		async (route) => {
			const line = (id: string, x1: number, y1: number, x2: number, y2: number) =>
				shape(
					id,
					cell('BeginX', x1) +
						cell('BeginY', y1) +
						cell('EndX', x2) +
						cell('EndY', y2) +
						(route ? cell('ShapeRouteStyle', 16) : '') +
						cell('ConLineJumpStyle', 3),
				);
			const glue = route
				? ''
				: '<Connects><Connect FromSheet="1" ToSheet="3" FromCell="BeginX" ToCell="PinX"/><Connect FromSheet="2" ToSheet="3" FromCell="BeginX" ToCell="PinX"/></Connects>';
			const source = await parseVsdx(
				await fixture({
					pages: [
						{
							id: '0',
							pageCells: cell('PageLineJumpDirX', 2),
							contents: `<Shapes>${line('1', 1, 2, 5, 2)}${line('2', 3, 0, 3, 4)}${shape('3')}</Shapes>${glue}`,
						},
					],
				}),
			);
			const before = [...visioLineJumpPaths(source.pages[0]!).values()].flat();
			expect(before).toHaveLength(1);
			const copy = copySnapshotScene(source);
			expect([...visioLineJumpPaths(copy.pages[0]!).values()].flat()).toEqual(before);
			expect(copy.pages[0]!.layout).not.toBe(source.pages[0]!.layout);
			const printed = await createPrintSnapshot(source, { pageIndices: [0] });
			expect(printed.pages[0]!.svg).toContain(`d="${before[0]}"`);
		},
	);
});

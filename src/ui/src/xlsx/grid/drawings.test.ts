import { afterEach, expect, it } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { createTestContext } from './test-context';
import { GridView } from './grid-view';
import { DrawingLayer } from './drawings';

afterEach(() => document.body.replaceChildren());
it.each(['pointercancel', 'lostpointercapture', 'Escape'])(
	'restores a drawing without editing its anchor on %s',
	(reason) => {
		const wb = createWorkbook();
		const anchor = {
			from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
			ext: { cx: 952500, cy: 476250 },
		};
		wb.sheets[0]!.drawings.push({ kind: 'unsupported', anchor, description: 'test shape' });
		const ctx = createTestContext(wb);
		const container = document.createElement('div');
		ctx.root.append(container);
		const view = new GridView(ctx, container);
		const layer = new DrawingLayer(view);
		layer.render();
		const node = container.querySelector<HTMLElement>('.xg-obj')!;
		const left = node.style.left;
		const pointer = (type: string, x: number) =>
			node.dispatchEvent(
				new PointerEvent(type, { pointerId: 1, button: 0, clientX: x, clientY: 20, bubbles: true }),
			);
		pointer('pointerdown', 20);
		pointer('pointermove', 60);
		expect(node.style.left).not.toBe(left);
		if (reason === 'Escape')
			document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		else pointer(reason, 60);
		pointer('pointerup', 60);
		expect(node.style.left).toBe(left);
		expect(ctx.workbook()!.sheets[0]!.drawings[0]!.anchor).toEqual(anchor);
		layer.destroy();
		view.destroy();
	},
);

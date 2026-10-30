// @vitest-environment jsdom
import { createWorkbook, putCell } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { mountGrid } from './index';
import { createTestContext } from './test-context';

let dispose: (() => void) | undefined;
afterEach(() => {
	dispose?.();
	dispose = undefined;
	document.body.replaceChildren();
});

const frame = () => new Promise((resolve) => setTimeout(resolve, 60));

describe('switching between a frozen and a split window', () => {
	it('keeps the scrolling pane where it was', async () => {
		const wb = createWorkbook();
		const ws = wb.sheets[0]!;
		for (let r = 0; r < 200; r++) putCell(ws, r, 0, { value: `r${r}` });
		ws.view.freeze = { rows: 1, cols: 1 };
		const ctx = createTestContext(wb);
		const container = document.createElement('div');
		ctx.root.append(container);
		dispose = mountGrid(ctx, container);
		const scroller = ctx.root.querySelector<HTMLElement>('.xg-scroller')!;
		const fire = async () => {
			scroller.dispatchEvent(new Event('scroll'));
			await frame();
		};
		scroller.scrollTop = 200;
		await fire();
		// Frozen scroll is relative to the frozen row (20 px): the pane's top is at 220.
		delete ws.view.freeze;
		ws.view.split = { x: 1500, y: 600 };
		await fire();
		expect(scroller.scrollTop).toBe(220);
		// And back: the scroll becomes relative to the frozen edge again.
		delete ws.view.split;
		ws.view.freeze = { rows: 1, cols: 1 };
		await fire();
		expect(scroller.scrollTop).toBe(200);
	});

	it('writes the scrolled first cells into the sheet view', async () => {
		const wb = createWorkbook();
		const ws = wb.sheets[0]!;
		ws.view.split = { x: 1500, y: 600 };
		const ctx = createTestContext(wb);
		const container = document.createElement('div');
		ctx.root.append(container);
		dispose = mountGrid(ctx, container);
		const scroller = ctx.root.querySelector<HTMLElement>('.xg-scroller')!;
		scroller.scrollTop = 200;
		scroller.dispatchEvent(new Event('scroll'));
		await frame();
		expect(ws.view.topLeft).toEqual({ row: 10, col: 0 });
		expect(ws.view.split?.topLeft).toEqual({ row: 0, col: 0 });
	});
});

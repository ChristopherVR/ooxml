// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { loadXlsx, type SparklineView } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { mountGrid } from './index';
import { sparklineSvg } from './sparkline';
import { createTestContext } from './test-context';

const fixture = () =>
	loadXlsx(
		new Uint8Array(
			readFileSync(resolve(process.cwd(), '../core/xlsx/__fixtures__/excel-sparklines.xlsx')),
		),
	);

let dispose: (() => void) | undefined;
afterEach(() => {
	dispose?.();
	dispose = undefined;
	document.body.replaceChildren();
});

const base: SparklineView = {
	type: 'line',
	color: '#376092',
	lineWeightPt: 0.75,
	lines: [],
	markers: [],
	columns: [],
};

const frame = () => new Promise((done) => setTimeout(done, 40));

describe('sparkline painter', () => {
	it('scales a line, its markers and the axis to the cell', () => {
		const svg = sparklineSvg(
			document,
			{
				...base,
				lines: [
					[
						{ x: 0, y: 1 },
						{ x: 1, y: 0 },
					],
				],
				markers: [{ x: 1, y: 0, color: '#D00000' }],
				axis: { y: 0.5, color: '#000000' },
			},
			100,
			20,
			100,
		);
		expect(svg.getAttribute('class')).toBe('xg-spark');
		expect(svg.getAttribute('width')).toBe('100');
		// 2px inset plus the 1.5px minimum marker radius on each side.
		expect(svg.querySelector('polyline')?.getAttribute('points')).toBe('3.5,16.5 96.5,3.5');
		expect(svg.querySelector('polyline')?.getAttribute('stroke')).toBe('#376092');
		const marker = svg.querySelector('circle')!;
		expect([marker.getAttribute('cx'), marker.getAttribute('fill')]).toEqual(['96.5', '#D00000']);
		expect(svg.querySelector('line')?.getAttribute('y1')).toBe('10');
		// Children paint in order: axis, columns, lines, markers.
		expect([...svg.children].map((c) => c.tagName)).toEqual(['line', 'polyline', 'circle']);
	});

	it('draws columns with a one-pixel sliver for values at the scale edge', () => {
		const svg = sparklineSvg(
			document,
			{
				...base,
				type: 'column',
				columns: [
					{ x: 0.05, y: 0, w: 0.4, h: 1, color: '#376092' },
					{ x: 0.55, y: 1, w: 0.4, h: 0, color: '#D00000' },
				],
			},
			104,
			24,
			100,
		);
		const rects = [...svg.querySelectorAll('rect')].map((r) =>
			['x', 'y', 'width', 'height', 'fill'].map((a) => r.getAttribute(a)),
		);
		expect(rects).toEqual([
			['7', '2', '40', '20', '#376092'],
			['57', '21', '40', '1', '#D00000'],
		]);
	});

	it('paints Excel sparklines in their host cells and repaints after edits and resizes', async () => {
		const workbook = await fixture();
		const ctx = createTestContext(workbook);
		const container = document.createElement('div');
		ctx.root.append(container);
		dispose = mountGrid(ctx, container);
		// The host cells F1:F4, top to bottom.
		const cell = (row: number) =>
			[...ctx.root.querySelectorAll<HTMLElement>('.xg-c')]
				.filter((n) => n.querySelector('.xg-spark'))
				.sort((a, b) => Number.parseFloat(a.style.top) - Number.parseFloat(b.style.top))[row];
		const sparks = ctx.root.querySelectorAll('.xg-spark');
		expect(sparks).toHaveLength(4);
		const line = cell(0)!;
		expect(line.firstElementChild?.getAttribute('class')).toBe('xg-spark');
		expect(line.querySelectorAll('circle')).toHaveLength(5);
		expect(cell(1)!.querySelectorAll('rect')).toHaveLength(4);
		expect(cell(2)!.querySelectorAll('rect')).toHaveLength(5);
		const before = line.querySelector('polyline')!.getAttribute('points');

		// Editing a cell in the data range repaints the sparkline.
		ctx.session()!.setCellValue(0, 0, 3, 50);
		await frame();
		const after = cell(0)!.querySelector('polyline')!.getAttribute('points');
		expect(after).not.toBe(before);

		// Widening the host column widens the sparkline with the cell.
		const width = Number(cell(0)!.querySelector('.xg-spark')!.getAttribute('width'));
		ctx.session()!.setColumnWidth(0, [5], 30);
		await frame();
		const host = cell(0)!;
		const wider = Number(host.querySelector('.xg-spark')!.getAttribute('width'));
		expect(wider).toBeGreaterThan(width);
		expect(wider).toBe(Number.parseFloat(host.style.width));
	});
});

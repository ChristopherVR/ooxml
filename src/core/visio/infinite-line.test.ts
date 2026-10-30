import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { geometryPaths, transform } from './geometry';
import { clipInfiniteLine, type InfiniteLineFrame } from './infinite-line';
import { readSheet } from './sheet';
import { cell, row, section, xml } from './test-fixtures';

const identity: InfiniteLineFrame = {
	toPage: [1, 0, 0, 1, 0, 0],
	width: 10,
	height: 6,
};

/** `null` means no page frame is supplied, which is distinct from the default identity frame. */
function infiniteLines(rows: string, frame: InfiniteLineFrame | null = identity) {
	const warnings: string[] = [];
	const sheet = readSheet(parseXml(xml('Shape', section('Geometry', rows))).documentElement);
	const paths = geometryPaths(
		sheet,
		frame?.width ?? 1,
		frame?.height ?? 1,
		(code) => warnings.push(code),
		() => {},
		() => {},
		frame ?? undefined,
	);
	return { paths, warnings };
}

const infiniteLine = (x: number, y: number, a: number, b: number) =>
	row(1, 'InfiniteLine', cell('X', x) + cell('Y', y) + cell('A', a) + cell('B', b));

describe('Visio InfiniteLine geometry', () => {
	it('draws a horizontal line across the full page width', () => {
		const { paths, warnings } = infiniteLines(infiniteLine(1, 2, 5, 2));
		expect(warnings).toEqual([]);
		expect(paths).toEqual([{ path: 'M 0 2 L 10 2', fill: false, stroke: true }]);
	});

	it('draws a vertical line across the full page height', () => {
		const { paths } = infiniteLines(infiniteLine(3, 1, 3, 4));
		expect(paths).toEqual([{ path: 'M 3 0 L 3 6', fill: false, stroke: true }]);
	});

	it('draws a diagonal line only where it crosses the page', () => {
		// y = x leaves the 10 x 6 page through the top edge at (6, 6).
		const { paths } = infiniteLines(infiniteLine(0, 0, 1, 1));
		expect(paths).toEqual([{ path: 'M 0 0 L 6 6', fill: false, stroke: true }]);
	});

	it('draws nothing for equal points', () => {
		const { paths, warnings } = infiniteLines(infiniteLine(2, 2, 2, 2));
		expect(paths).toEqual([]);
		expect(warnings).toEqual([]);
	});

	it('draws nothing for a line that misses the page', () => {
		const { paths, warnings } = infiniteLines(infiniteLine(0, 10, 1, 10));
		expect(paths).toEqual([]);
		expect(warnings).toEqual([]);
	});

	it('clips a line through a rotated, translated shape to the page', () => {
		// Pin (5, 3), local pin (2, 1), rotated 90 degrees: page = (6 - y, 1 + x).
		const toPage = transform(5, 3, 2, 1, Math.PI / 2);
		const { paths } = infiniteLines(infiniteLine(0, 0, 1, 0), { toPage, width: 10, height: 6 });
		// The local x-axis maps to the page line x = 6, which is visible for local x in [-1, 5].
		expect(paths).toEqual([{ path: 'M -1 0 L 5 0', fill: false, stroke: true }]);
	});

	it('honours the section NoLine cell and never fills the line', () => {
		const { paths } = infiniteLines(cell('NoLine', 1) + infiniteLine(1, 2, 5, 2));
		expect(paths).toEqual([{ path: 'M 0 2 L 10 2', fill: false, stroke: false }]);
	});

	it('reports the line as unsupported when no page frame is available', () => {
		const { paths, warnings } = infiniteLines(infiniteLine(0, 0, 1, 1), null);
		expect(paths).toEqual([]);
		expect(warnings).toEqual(['unsupported-geometry']);
	});
});

describe('clipInfiniteLine', () => {
	it('returns the same parameter in local space as in page space', () => {
		const frame: InfiniteLineFrame = { toPage: [2, 0, 0, 2, 1, 1], width: 21, height: 11 };
		// Page points are (1, 1) + 2 * local point, so page x in [0, 21] is local x in [-0.5, 10].
		expect(clipInfiniteLine(0, 0, 1, 0, frame)).toEqual([-0.5, 0, 10, 0]);
	});
});

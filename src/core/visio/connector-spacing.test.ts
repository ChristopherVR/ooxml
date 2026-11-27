import { describe, expect, it } from 'vitest';
import {
	VISIO_CONNECTOR_SPACING_DEFAULTS,
	spaceVisioRoute,
	visioConnectorSpacing,
	type VisioSpacedRoute,
} from './connector-spacing';

const p = (x: number, y: number) => ({ x, y });
/** A Z route from (0, 0) to (4, 2) whose middle run is vertical at x = 2. */
const zed = [p(0, 0), p(2, 0), p(2, 2), p(4, 2)];
const other = (points: { x: number; y: number }[], ...targets: string[]): VisioSpacedRoute => ({
	points,
	targets: new Set(targets),
});
const defaults = VISIO_CONNECTOR_SPACING_DEFAULTS;

describe('parallel connector spacing', () => {
	it('reads LineAdjustFrom as Page Setup > Layout and Routing names it', () => {
		expect(visioConnectorSpacing(undefined)).toEqual(defaults);
		expect(visioConnectorSpacing(0).separate).toBe('unrelated');
		expect(visioConnectorSpacing(1).separate).toBe('all');
		expect(visioConnectorSpacing(2).separate).toBe('none');
		expect(visioConnectorSpacing(3).separate).toBe('unrelated');
		expect(visioConnectorSpacing(0, 0.25, -1)).toMatchObject({ lineToLineX: 0.25, lineToLineY: 0 });
	});

	it('moves an interior run that lies on an unrelated connector by one line spacing', () => {
		const spaced = spaceVisioRoute(
			zed,
			new Set(['A', 'B']),
			[other([p(2, -1), p(2, 3)], 'C', 'D')],
			defaults,
		);
		expect(spaced).toEqual([p(0, 0), p(2.125, 0), p(2.125, 2), p(4, 2)]);
	});

	it('keeps the ends where they are glued', () => {
		const spaced = spaceVisioRoute(zed, new Set(), [other([p(2, -1), p(2, 3)])], defaults);
		expect(spaced[0]).toEqual(zed[0]);
		expect(spaced.at(-1)).toEqual(zed.at(-1));
	});

	it('steps further out while the nearer place is also taken', () => {
		const spaced = spaceVisioRoute(
			zed,
			new Set(),
			[other([p(2, -1), p(2, 3)]), other([p(2.125, -1), p(2.125, 3)])],
			defaults,
		);
		// +1/8 is taken, so the other side, -1/8, is next.
		expect(spaced[1]).toEqual(p(1.875, 0));
	});

	it('lets related connectors (sharing a glued shape) overlap unless all lines are separated', () => {
		const shared = [other([p(2, -1), p(2, 3)], 'A')];
		expect(spaceVisioRoute(zed, new Set(['A', 'B']), shared, defaults)).toEqual(zed);
		expect(spaceVisioRoute(zed, new Set(['A', 'B']), shared, visioConnectorSpacing(1))[1]).toEqual(
			p(2.125, 0),
		);
	});

	it('does nothing when the page separates no lines or the spacing is zero', () => {
		const blocker = [other([p(2, -1), p(2, 3)])];
		expect(spaceVisioRoute(zed, new Set(), blocker, visioConnectorSpacing(2))).toEqual(zed);
		expect(spaceVisioRoute(zed, new Set(), blocker, visioConnectorSpacing(0, 0, 0))).toEqual(zed);
	});

	it('leaves runs that only touch, cross or are parallel at a distance', () => {
		const touching = [other([p(2, 2), p(2, 5)])];
		const crossing = [other([p(1, 1), p(3, 1)])];
		const apart = [other([p(2.5, -1), p(2.5, 3)])];
		for (const others of [touching, crossing, apart])
			expect(spaceVisioRoute(zed, new Set(), others, defaults)).toEqual(zed);
	});

	it('never moves a run so far that a neighbouring run turns back into its shape', () => {
		// The first run is 0.1 long, so the middle run may only move right.
		const short = [p(0, 0), p(0.1, 0), p(0.1, 2), p(4, 2)];
		const spaced = spaceVisioRoute(short, new Set(), [other([p(0.1, -1), p(0.1, 3)])], defaults);
		expect(spaced[1]).toEqual(p(0.225, 0));
	});

	it('does not move a run into an obstacle', () => {
		const spaced = spaceVisioRoute(zed, new Set(), [other([p(2, -1), p(2, 3)])], defaults, [
			{ minX: 2.05, minY: 0.5, maxX: 2.2, maxY: 1.5 },
		]);
		expect(spaced[1]).toEqual(p(1.875, 0));
	});

	it('leaves one- and two-run routes alone: their runs leave or enter a shape', () => {
		const straight = [p(0, 0), p(4, 0)];
		const ell = [p(0, 0), p(2, 0), p(2, 2)];
		expect(spaceVisioRoute(straight, new Set(), [other([p(-1, 0), p(5, 0)])], defaults)).toEqual(
			straight,
		);
		expect(spaceVisioRoute(ell, new Set(), [other([p(-1, 0), p(5, 0)])], defaults)).toEqual(ell);
	});
});

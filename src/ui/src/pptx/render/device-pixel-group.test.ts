import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { connectorSnapTransform } from './connector-path';
import { snapPlacement } from './device-pixel-group';
import { getGroupChildParentFill } from './group-fill';

function group(over: Record<string, unknown>, children: PptxElement[]): PptxElement {
	return { type: 'group', id: 'g', children, ...over } as unknown as PptxElement;
}

function line(over: Record<string, unknown> = {}): PptxElement {
	return { type: 'connector', id: 'c', x: 10.4, y: 20.4, width: 50, height: 0, ...over } as never;
}

describe('snapPlacement', () => {
	it('places a top-level element at its own position', () => {
		const el = line();
		expect(snapPlacement(el)).toMatchObject({ x: 10.4, y: 20.4 });
	});

	it('adds the ancestor groups, painted and exact, once the group is registered', () => {
		const child = line();
		// The group origin is rounded in the model (100.3 px) but exact in EMU.
		const inner = group({ x: 5.25, y: 3, width: 80, height: 40 }, [child]);
		const outer = group(
			{ x: 100, y: 200, xEmu: Math.round(100.3 * 9525), yEmu: 200 * 9525, width: 99, height: 99 },
			[inner],
		);
		getGroupChildParentFill(outer);
		getGroupChildParentFill(inner);
		const place = snapPlacement(child);
		// The painted position is the whole stage pixel Chromium snaps the svg root to.
		expect(place?.x).toBe(116);
		expect(place?.y).toBe(223);
		// The exact box starts at the unrounded 100.3 origin, not the painted 100.
		expect(place?.exact.x).toBeCloseTo(100.3 + 5.25 + 10.4, 3);
	});

	it('does not snap inside a rotated or flipped group', () => {
		const child = line();
		const rotated = group({ x: 0, y: 0, rotation: 30 }, [child]);
		getGroupChildParentFill(rotated);
		expect(snapPlacement(child)).toBeUndefined();
		const flipped = line();
		getGroupChildParentFill(group({ x: 0, y: 0, flipVertical: true }, [flipped]));
		expect(snapPlacement(flipped)).toBeUndefined();
	});
});

describe('group child strokes', () => {
	it('snaps a line in a group by its absolute position, not its group-relative one', () => {
		const child = line({ id: 'child', x: 16.4, y: 23.4 });
		const alone = { ...child } as PptxElement;
		getGroupChildParentFill(group({ x: 100, y: 200, width: 200, height: 100 }, [child]));
		const inside = connectorSnapTransform(child, true, 1, 1);
		const outside = connectorSnapTransform(alone, true, 1, 1);
		expect(inside).toBeDefined();
		// 16.4 px and 116.4 px round to different device pixels at fractional
		// zoom, and the painted position differs: the translate must differ.
		expect(inside).not.toBe(outside);
		expect(inside).toContain('round(223.4px');
	});

	it('leaves a line in a rotated group unsnapped', () => {
		const child = line();
		getGroupChildParentFill(group({ x: 0, y: 0, rotation: 90 }, [child]));
		expect(connectorSnapTransform(child, true, 1, 1)).toBeUndefined();
	});
});

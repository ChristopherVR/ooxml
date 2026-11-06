import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { bodyTopOf, bodyTopSnapShift } from './text-body-top-snap';

const textBox = (over: Record<string, unknown> = {}, style: Record<string, unknown> = {}) =>
	({
		type: 'text',
		id: 't',
		x: 11,
		y: 13,
		width: 80,
		height: 200,
		text: 'H',
		textSegments: [{ text: 'H', style: {} }],
		textStyle: { bodyInsetTop: 0, ...style },
		...over,
	}) as unknown as PptxElement;

describe('bodyTopOf', () => {
	it('is the shape top plus the top inset, painted and exact', () => {
		// The model rounds a 13.33px shape top to 13; the real position is the EMU one.
		const top = bodyTopOf(textBox({ yEmu: 127000 }, { bodyInsetTop: 4.8 }));
		expect(top?.painted).toBeCloseTo(17.8, 5);
		expect(top?.exact).toBeCloseTo(13.333 + 4.8, 2);
	});

	it('does not apply to other anchors, rotated shapes or vertical text', () => {
		expect(bodyTopOf(textBox({}, { vAlign: 'middle' }))).toBeUndefined();
		expect(bodyTopOf(textBox({}, { vAlign: 'bottom' }))).toBeUndefined();
		expect(bodyTopOf(textBox({ rotation: 15 }))).toBeUndefined();
		expect(bodyTopOf(textBox({}, { textDirection: 'vertical' }))).toBeUndefined();
	});
});

describe('bodyTopSnapShift', () => {
	// PowerPoint 16, shape top swept in 0.05pt steps: the first baseline moved
	// down by a pixel when the body top crossed .5 of a pixel (12 to 22pt).
	it('moves the body top onto the pixel PowerPoint snaps to', () => {
		expect(bodyTopSnapShift({ painted: 13, exact: 13.33 }, 16)).toBe(0);
		expect(bodyTopSnapShift({ painted: 14, exact: 13.33 }, 16)).toBe(-1);
		expect(bodyTopSnapShift({ painted: 13, exact: 13.67 }, 16)).toBe(1);
	});

	it('snaps larger text a third of a pixel earlier (24pt and up)', () => {
		expect(bodyTopSnapShift({ painted: 13, exact: 13.33 }, 32)).toBe(1);
		expect(bodyTopSnapShift({ painted: 13, exact: 13.33 }, 31)).toBe(0);
	});
});

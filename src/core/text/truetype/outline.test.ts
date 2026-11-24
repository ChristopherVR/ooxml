import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { glyphOutlinePath } from './outline';
import { parseTrueType } from './parse';
import { outlineTrueType } from './test-support/outline-font';

describe('glyphOutlinePath', () => {
	const font = parseTrueType(outlineTrueType())!;

	it('reads a simple glyph of straight lines', () => {
		expect(glyphOutlinePath(font, font.glyphFor(0x41))).toEqual([
			{ t: 'M', x: 0, y: 0 },
			{ t: 'L', x: 0, y: 700 },
			{ t: 'L', x: 600, y: 700 },
			{ t: 'L', x: 600, y: 0 },
			{ t: 'L', x: 0, y: 0 },
			{ t: 'Z' },
		]);
	});

	it('puts implied on-curve points between off-curve points', () => {
		const path = glyphOutlinePath(font, font.glyphFor(0x42));
		expect(path[0]).toEqual({ t: 'M', x: 450, y: 175 });
		expect(path.filter((c) => c.t === 'Q')).toHaveLength(4);
		expect(path[1]).toEqual({ t: 'Q', cx: 300, cy: 0, x: 150, y: 175 });
		expect(path.at(-2)).toEqual({ t: 'Q', cx: 600, cy: 350, x: 450, y: 175 });
	});

	it('places and scales the components of a composite glyph', () => {
		const points = glyphOutlinePath(font, font.glyphFor(0x43)).flatMap((c) =>
			c.t === 'M' || c.t === 'L' ? [[c.x, c.y]] : [],
		);
		expect(points).toEqual([
			[100, 0],
			[100, 350],
			[400, 350],
			[400, 0],
			[100, 0],
		]);
	});

	it('is empty for blank glyphs and missing ones', () => {
		expect(glyphOutlinePath(font, font.glyphFor(0x20))).toEqual([]);
		expect(glyphOutlinePath(font, 999)).toEqual([]);
	});

	const arial = 'C:/Windows/Fonts/arial.ttf';
	it.skipIf(!existsSync(arial))('reads the closed contours of a real font', () => {
		const real = parseTrueType(new Uint8Array(readFileSync(arial)))!;
		const path = glyphOutlinePath(real, real.glyphFor(0x4f));
		// O: an outer and an inner contour.
		expect(path.filter((c) => c.t === 'M')).toHaveLength(2);
		expect(path.filter((c) => c.t === 'Z')).toHaveLength(2);
		for (const c of path)
			if (c.t !== 'Z') expect(Math.abs(c.x) + Math.abs(c.y)).toBeLessThan(4 * real.unitsPerEm);
	});
});

import { describe, expect, it } from 'vitest';
import { outlineTrueType } from '../../text/truetype/test-support/outline-font';
import { PdfEmbeddedFont } from './pdf-embedded-font';
import { glyphSegments, placeGlyphs, runOutlines } from './pdf-glyph-outlines';
import { parsePath } from './pdf-path';
import { pathSampler } from './pdf-path-sampler';

const font = PdfEmbeddedFont.create({ family: 'Outline', data: outlineTrueType() })!;

describe('glyph outlines', () => {
	it('places glyphs as embeddedText writes them: spacing and word spacing, scaled by Tz', () => {
		const placed = placeGlyphs(font, 10, 'A A', { character: 1, word: 2, scale: 50 });
		// A: (7 + 1) x 0.5; space: (3 + 1 + 2) x 0.5.
		expect(placed.map((g) => [g.char, g.x, g.advance])).toEqual([
			['A', 0, 4],
			[' ', 4, 3],
			['A', 7, 4],
		]);
	});

	it('writes a straight run as outlines in y-down page space', () => {
		const segments = runOutlines(font, 10, 100, 50, 'A');
		expect(segments.slice(0, 4)).toEqual([
			{ t: 'M', x: 100, y: 50 },
			{ t: 'L', x: 100, y: 43 },
			{ t: 'L', x: 106, y: 43 },
			{ t: 'L', x: 106, y: 50 },
		]);
		// The off-curve ring becomes cubic curves.
		expect(runOutlines(font, 10, 0, 0, 'B').filter((s) => s.t === 'C')).toHaveLength(4);
	});

	it('bends an outline along a curved map, cutting its straight edges', () => {
		const arc = pathSampler(parsePath('M 0 100 A 100 100 0 0 1 200 100')!)!;
		const map = (u: number, v: number) => {
			const p = arc.at(u);
			return [p.x - Math.sin(p.angle) * v, p.y + Math.cos(p.angle) * v] as const;
		};
		const bent = glyphSegments(font, font.font.glyphFor(0x41), 100, 100, map, { bend: true });
		expect(bent.filter((s) => s.t === 'L').length).toBeGreaterThan(20);
		// Every point of the bent square lies between the arc and 70 units outside it.
		for (const s of bent)
			if (s.t === 'L') {
				const r = Math.hypot(s.x - 100, s.y - 100);
				expect(r).toBeGreaterThan(99.5);
				expect(r).toBeLessThan(170.5);
			}
	});
});

describe('pathSampler', () => {
	it('walks a path by length and carries on straight past its ends', () => {
		const sampler = pathSampler(parsePath('M 0 0 L 100 0 L 100 50')!)!;
		expect(sampler.length).toBeCloseTo(150);
		expect(sampler.at(50)).toMatchObject({ x: 50, y: 0, angle: 0 });
		expect(sampler.at(125).y).toBeCloseTo(25);
		expect(sampler.at(125).angle).toBeCloseTo(Math.PI / 2);
		expect(sampler.at(-10)).toMatchObject({ x: -10, y: 0 });
		expect(sampler.at(160).y).toBeCloseTo(60);
		expect(pathSampler(parsePath('M 5 5')!)).toBeUndefined();
	});
});

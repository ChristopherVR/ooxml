import { describe, expect, it } from 'vitest';
import { buildWarpPath } from './text-warp-path';
import { wordArtLayout, wordArtShadowGeometry } from './word-art-layout';
import { renderWordArtSvg, wordArtCssShadow } from './word-art-svg';

describe('wordArtLayout', () => {
	it('fits one size and centres straight lines a line pitch apart', () => {
		const layout = wordArtLayout({
			style: { fill: '112233', outline: '445566' },
			lines: ['One', 'Two'],
			width: 300,
			height: 200,
			fontPx: 40,
		})!;
		expect(layout.size).toBe(40);
		expect(layout.preset).toBeUndefined();
		expect(layout.outline).toEqual({ color: '445566', width: 4 / 3 });
		expect(layout.lines).toEqual([
			{ text: 'One', kind: 'straight', x: 150, y: 100 + 40 * 0.34 - 23 },
			{ text: 'Two', kind: 'straight', x: 150, y: 100 + 40 * 0.34 + 23 },
		]);
	});

	it('puts warped lines on the preset path inside the inset box, as the SVG preview does', () => {
		const style = { fill: '112233', warp: { preset: 'textArchUp', guides: [] } };
		const layout = wordArtLayout({ style, lines: ['Arc'], width: 300, height: 200, fontPx: 40 })!;
		const [line] = layout.lines;
		expect(line).toEqual({
			text: 'Arc',
			kind: 'path',
			path: buildWarpPath('textArchUp', 260, 200 - 32 - 40 * 0.26, 0, 1, undefined, undefined, 46),
			offsetX: 20,
			offsetY: 32,
		});
		const svg = renderWordArtSvg({
			id: 'w',
			style,
			lines: ['Arc'],
			width: 300,
			height: 200,
			fontPx: 40,
		});
		expect(svg).toContain(line!.kind === 'path' ? line!.path : '');
		expect(svg).toContain('translate(20 32)');
	});

	it('keeps the shadow the CSS preview draws', () => {
		const shadow = wordArtShadowGeometry({ fill: '000000', shadow: '45460D' })!;
		expect(shadow).toMatchObject({ color: '45460D', blur: 8 / 3, opacity: 0.55 });
		expect(shadow.dx).toBeCloseTo(Math.SQRT1_2 * 4);
		expect(wordArtCssShadow({ fill: '000000', shadow: '45460D' })).toBe(
			`${shadow.dx}px ${shadow.dy}px ${shadow.blur}px #45460D8c`,
		);
		expect(wordArtShadowGeometry({ fill: '000000' })).toBeUndefined();
	});

	it('draws nothing in an empty box', () => {
		expect(
			wordArtLayout({ style: { fill: '000000' }, lines: ['x'], width: 0, height: 9, fontPx: 9 }),
		).toBe(undefined);
	});
});

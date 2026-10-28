import { expect, it } from 'vitest';
import { picturePaint } from './picture-paint';
import { pictureEffectFilter } from './picture-effects';
const defs = () => document.createElementNS('http://www.w3.org/2000/svg', 'defs');
it('paints native-size mirrored tiles with crop and page-fixed orientation', () => {
	const target = defs();
	const result = picturePaint(
		document,
		target,
		{
			kind: 'picture',
			crop: { cropLeft: 0.25 },
			layout: {
				tile: { scaleX: 2, scaleY: 1, alignment: 'ctr', offsetXEmu: 9525, flip: 'xy' },
				rotWithShape: false,
			},
		},
		{ url: 'data:image/png;base64,AA==', width: 20, height: 10 },
		{ width: 100, height: 50, rotation: 30, flipHorizontal: true },
	);
	expect(result.unapplied).toEqual([]);
	const pattern = target.querySelector('pattern')!;
	expect(pattern.getAttribute('width')).toBe('80');
	expect(pattern.getAttribute('height')).toBe('20');
	expect(pattern.getAttribute('x')).toBe('31');
	expect(pattern.querySelectorAll('image')).toHaveLength(4);
	expect(Number(pattern.querySelector('image')!.getAttribute('width'))).toBeCloseTo(40 / 0.75);
	expect(pattern.getAttribute('patternTransform')).toContain('rotate(-30)');
});
it('clips a stretched cropped image to its offset fillRect', () => {
	const target = defs();
	picturePaint(
		document,
		target,
		{
			kind: 'picture',
			crop: { cropLeft: 0.25, cropRight: 0.25 },
			layout: { fillRect: { l: 0.1, t: 0.2, r: 0.3, b: 0.1 } },
		},
		'data:image/png;base64,AA==',
		{ width: 100, height: 50 },
	);
	expect(target.querySelector('image')?.getAttribute('width')).toBe('1.2000000000000002');
	expect(target.querySelector('clipPath rect')?.getAttribute('x')).toBe('0.1');
});
it('chains richer alpha, duotone and artistic preview primitives without applying baked art twice', () => {
	const target = defs();
	const effect = pictureEffectFilter(document, target, [
		{ kind: 'standard', name: 'alphaInv' },
		{
			kind: 'standard',
			name: 'duotone',
			colors: [
				{ kind: 'srgb', value: '000000', transforms: [] },
				{ kind: 'srgb', value: 'FF0000', transforms: [] },
			],
		},
		{ kind: 'artistic', name: 'artisticFilmGrain', radius: 30, baked: false },
	]);
	expect(effect.unapplied).toEqual([]);
	expect(effect.approximated).toEqual(['artisticFilmGrain']);
	expect(target.querySelectorAll('feTurbulence')).toHaveLength(1);
	expect(target.querySelector('feFuncA')?.getAttribute('slope')).toBe('-1');
	const baked = defs();
	expect(
		pictureEffectFilter(document, baked, [
			{ kind: 'artistic', name: 'artisticFilmGrain', radius: 30, baked: true },
		]).filter,
	).toBeUndefined();
});

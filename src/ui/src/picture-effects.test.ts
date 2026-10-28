import { expect, it } from 'vitest';
import { pictureEffectFilter } from './picture-effects';
it('paints supported effects in source order and reports unavailable operations', () => {
	const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
	const result = pictureEffectFilter(document, defs, [
		{ kind: 'alphaModFix', amount: 0.5 },
		{ kind: 'grayscale' },
		{ kind: 'luminance', brightness: 0.7, contrast: -0.7 },
		{ kind: 'unsupported', name: 'future' },
	]);
	expect(result.unapplied).toEqual(['future']);
	expect(result.filter).toMatch(/^url/);
	expect([...defs.firstElementChild!.children].map((node) => node.localName)).toEqual([
		'feComponentTransfer',
		'feColorMatrix',
		'feComponentTransfer',
	]);
	expect(defs.querySelector('feFuncA')?.getAttribute('slope')).toBe('0.5');
	expect(defs.querySelector('feFuncR')?.getAttribute('intercept')).toBe('0.805');
});

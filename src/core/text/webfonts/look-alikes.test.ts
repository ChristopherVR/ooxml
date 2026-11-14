import { describe, expect, it } from 'vitest';

import { matchGoogleWebfontFragments, servedGoogleFontsFamily } from './google-fonts-request';
import { findLookAlikeGoogleFontsFamily, findLookAlikeWeights } from './look-alikes';
import { collectWebfontFaces } from './webfont-faces';

describe('findLookAlikeGoogleFontsFamily', () => {
	it('maps Aptos and Aptos Display to their measured look-alike, in any case', () => {
		expect(findLookAlikeGoogleFontsFamily('Aptos')).toBe('Gantari');
		expect(findLookAlikeGoogleFontsFamily(' aptos display ')).toBe('Gantari');
	});

	it('has nothing for a served family or one without a measured look-alike', () => {
		expect(findLookAlikeGoogleFontsFamily('Roboto')).toBeNull();
		expect(findLookAlikeGoogleFontsFamily('Calibri')).toBeNull();
		expect(findLookAlikeGoogleFontsFamily('Aptos Narrow')).toBeNull();
	});

	it('gives the weights that match the font, only for a look-alike', () => {
		expect(findLookAlikeWeights('Aptos')).toStrictEqual({ regular: 400, bold: 650 });
		expect(findLookAlikeWeights('Consolas')).toStrictEqual({ regular: 500, bold: 700 });
		expect(findLookAlikeWeights('Calibri')).toBeUndefined();
	});

	it('maps the measured Office fonts to their look-alikes', () => {
		expect(findLookAlikeGoogleFontsFamily('Comic Sans MS')).toBe('Comic Relief');
		expect(findLookAlikeGoogleFontsFamily('Consolas')).toBe('Overpass Mono');
		expect(findLookAlikeGoogleFontsFamily('Aptos Mono')).toBe('Sometype Mono');
		expect(findLookAlikeGoogleFontsFamily('Garamond')).toBe('EB Garamond');
		expect(findLookAlikeGoogleFontsFamily('Impact')).toBe('Anton');
		expect(findLookAlikeGoogleFontsFamily('Franklin Gothic Medium')).toBeNull();
	});
});

describe('look-alikes are opt-in', () => {
	it('keeps the metric clone unless the product corrects widths itself', () => {
		expect(servedGoogleFontsFamily('Aptos')).toBe('Carlito');
		expect(servedGoogleFontsFamily('Aptos', { lookAlikes: true })).toBe('Gantari');
		expect(servedGoogleFontsFamily('Calibri', { lookAlikes: true })).toBe('Carlito');
	});

	it('requests and waits for the look-alike only when asked', () => {
		expect(matchGoogleWebfontFragments(['Aptos'])[0]).toMatch(/^Carlito:/u);
		expect(matchGoogleWebfontFragments(['Aptos'], undefined, { lookAlikes: true })).toStrictEqual([
			'Gantari:ital,wght@0,400;0,650;1,400;1,650',
		]);
		// The Garamond look-alike was already listed in the CSS chain but never requested.
		expect(
			matchGoogleWebfontFragments(['Garamond', 'Impact'], undefined, { lookAlikes: true }),
		).toStrictEqual([
			'Anton:ital,wght@0,400;1,400',
			'EB Garamond:ital,wght@0,400;0,600;1,400;1,600',
		]);
		expect(matchGoogleWebfontFragments(['Garamond'])).toStrictEqual([]);
		const runs = [{ family: 'Aptos', bold: true, text: 'Hi' }];
		const { faces } = collectWebfontFaces(runs, ['Aptos'], { lookAlikes: true });
		expect([...faces.keys()]).toStrictEqual(['Gantari']);
		expect(faces.get('Gantari')?.[0]?.font).toBe('700 16px "Gantari"');
	});
});

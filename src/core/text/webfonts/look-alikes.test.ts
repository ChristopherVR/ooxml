import { describe, expect, it } from 'vitest';

import { matchGoogleWebfontFragments, servedGoogleFontsFamily } from './google-fonts-request';
import { findLookAlikeGoogleFontsFamily } from './look-alikes';
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
});

describe('look-alikes are opt-in', () => {
	it('keeps the metric clone unless the product corrects widths itself', () => {
		expect(servedGoogleFontsFamily('Aptos')).toBe('Carlito');
		expect(servedGoogleFontsFamily('Aptos', { lookAlikes: true })).toBe('Gantari');
		expect(servedGoogleFontsFamily('Calibri', { lookAlikes: true })).toBe('Carlito');
	});

	it('requests and waits for the look-alike only when asked', () => {
		expect(matchGoogleWebfontFragments(['Aptos'])[0]).toMatch(/^Carlito:/u);
		expect(matchGoogleWebfontFragments(['Aptos'], undefined, { lookAlikes: true })[0]).toMatch(
			/^Gantari:/u,
		);
		const runs = [{ family: 'Aptos', bold: true, text: 'Hi' }];
		const { faces } = collectWebfontFaces(runs, ['Aptos'], { lookAlikes: true });
		expect([...faces.keys()]).toStrictEqual(['Gantari']);
		expect(faces.get('Gantari')?.[0]?.font).toBe('700 16px "Gantari"');
	});
});

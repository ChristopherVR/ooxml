import { describe, expect, it, vi } from 'vitest';

import {
	GOOGLE_FONTS_CSS2_BASE,
	buildGoogleFontsHref,
	collectRunFamilies,
	collectWebfontFaces,
	matchGoogleWebfontFragments,
	resolveGoogleFontsCss2Base,
	servedGoogleFontsFamily,
} from './index';
import { metricFontFamily } from '../metric-font-family';

describe('shared webfont request', () => {
	it('lets each product choose what an unset option means', () => {
		expect(resolveGoogleFontsCss2Base(undefined)).toBe(GOOGLE_FONTS_CSS2_BASE);
		expect(resolveGoogleFontsCss2Base(undefined, null, 'docx-editor')).toBeNull();
		expect(resolveGoogleFontsCss2Base(false, GOOGLE_FONTS_CSS2_BASE)).toBeNull();
		expect(resolveGoogleFontsCss2Base({ baseUrl: 'https://fonts.example.com/css2' }, null)).toBe(
			'https://fonts.example.com/css2',
		);
	});

	it('fails closed on an invalid mirror, warning once with the product label', () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
		try {
			expect(resolveGoogleFontsCss2Base({ baseUrl: 'not a url ' + Math.random() }, null, 'x')).toBe(
				null,
			);
			expect(warn).toHaveBeenCalledWith(expect.stringContaining('[x] googleFonts.baseUrl'));
		} finally {
			warn.mockRestore();
		}
	});

	it('requests a catalogue family as itself and an Office font as its metric clone', () => {
		expect(servedGoogleFontsFamily('roboto')).toBe('Roboto');
		expect(servedGoogleFontsFamily('Calibri')).toBe('Carlito');
		expect(servedGoogleFontsFamily('Aptos')).toBe('Carlito');
		expect(servedGoogleFontsFamily('Contoso Corporate Sans')).toBeNull();
		const resolved: string[] = [];
		const fragments = matchGoogleWebfontFragments(['Calibri', 'Aptos', 'Private Font'], (f) =>
			resolved.push(f),
		);
		expect(resolved).toEqual(['Calibri', 'Aptos']);
		expect(buildGoogleFontsHref(fragments)).toBe(
			'https://fonts.googleapis.com/css2?family=Carlito%3Aital%2Cwght%400%2C400%3B0%2C700%3B1%2C400%3B1%2C700&display=swap',
		);
	});

	it('lists the clone it requests in the shared CSS stack, so the download takes effect', () => {
		for (const family of ['Calibri', 'Aptos', 'Cambria', 'Arial', 'Times New Roman']) {
			const clone = servedGoogleFontsFamily(family);
			expect(clone).not.toBeNull();
			expect(metricFontFamily(family)).toContain(`"${clone}"`);
		}
	});

	it('collects faces and characters per served family from neutral runs', () => {
		const runs = [
			{ family: 'Calibri', text: 'Hello' },
			{ family: 'Calibri', bold: true, text: 'World' },
			{ family: 'Private Font', text: 'x' },
			{ family: undefined, text: 'none' },
		];
		expect([...collectRunFamilies(runs)]).toEqual(['Calibri', 'Private Font']);
		const { faces, unavailable } = collectWebfontFaces(runs, ['Calibri', 'Private Font']);
		expect(unavailable).toEqual(['Private Font']);
		const carlito = faces.get('Carlito') ?? [];
		expect(carlito.map((face) => face.font)).toEqual(['400 16px "Carlito"', '700 16px "Carlito"']);
		expect([...(carlito[0]?.chars ?? [])].join('')).toBe('Helo');
	});
});

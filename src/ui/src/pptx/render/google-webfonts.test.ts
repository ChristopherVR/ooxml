import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { GOOGLE_FONTS_CATALOGUE_DATE, GOOGLE_FONTS_FAMILIES } from './google-fonts-catalogue';
import {
	buildGoogleFontsFragment,
	buildGoogleFontsHref,
	collectReferencedFontFamilies,
	isFontFamilyInstalledLocally,
	findGoogleFontsFamily,
	findMetricCompatibleGoogleFontsFamily,
	GOOGLE_FONTS_CSS2_BASE,
	matchGoogleWebfontFragments,
	normalizeGoogleFontsBaseUrl,
	resetGoogleWebfontSessionCache,
	resolveGoogleFontsCss2Base,
	resolveGoogleWebfontHref,
	selectGoogleWebfontFamilies,
} from './google-webfonts';

/** Minimal text element carrying one styled segment. */
function textEl(fontFamily?: string): PptxElement {
	return {
		type: 'text',
		textSegments: fontFamily ? [{ style: { fontFamily } }] : [{ style: {} }],
	} as unknown as PptxElement;
}

function slide(...elements: PptxElement[]): PptxSlide {
	return { elements } as unknown as PptxSlide;
}

beforeEach(() => {
	resetGoogleWebfontSessionCache();
});

afterEach(() => {
	vi.restoreAllMocks();
	// The stubbed fetch must not outlive this file: the next file in the worker would load decks
	// with it, and a failed blob fetch leaves random blob URLs in their view models.
	vi.unstubAllGlobals();
});

describe('collectReferencedFontFamilies', () => {
	it('collects unique families across slides and dedupes', () => {
		const families = collectReferencedFontFamilies([
			slide(textEl('ADLaM Display'), textEl('Arial'), textEl('ADLaM Display')),
			slide(textEl('Roboto')),
		]);
		expect([...families]).toStrictEqual(['ADLaM Display', 'Arial', 'Roboto']);
	});

	it('recurses into group children', () => {
		const group = {
			type: 'group',
			children: [textEl('ADLaM Display')],
		} as unknown as PptxElement;
		expect([...collectReferencedFontFamilies([slide(group)])]).toStrictEqual(['ADLaM Display']);
	});

	it('counts an element-level fontFamily on segment-less text', () => {
		const bodyStyled = {
			type: 'text',
			textStyle: { fontFamily: 'Roboto' },
		} as unknown as PptxElement;
		expect([...collectReferencedFontFamilies([slide(bodyStyled)])]).toStrictEqual(['Roboto']);
	});

	it('skips segments without a family and empty decks', () => {
		expect(collectReferencedFontFamilies([slide(textEl())]).size).toBe(0);
		expect(collectReferencedFontFamilies([]).size).toBe(0);
	});
});

describe('selectGoogleWebfontFamilies', () => {
	it('keeps every referenced family the deck does not embed', () => {
		expect(
			selectGoogleWebfontFamilies(['ADLaM Display', 'Roboto', 'Some Local Font'], ['Roboto']),
		).toStrictEqual(['ADLaM Display', 'Some Local Font']);
	});

	it('accepts an empty embedded list', () => {
		expect(selectGoogleWebfontFamilies(['Cabin'], [])).toStrictEqual(['Cabin']);
	});

	it('drops families the runtime reports as installed before anything is requested', () => {
		expect(
			selectGoogleWebfontFamilies(
				['Installed Face', 'Missing Face'],
				[],
				(family) => family === 'Installed Face',
			),
		).toStrictEqual(['Missing Face']);
	});
});

describe('isFontFamilyInstalledLocally', () => {
	it('returns false when there is no DOM to ask (pure runtime)', () => {
		expect(isFontFamilyInstalledLocally('Arial')).toBeFalsy();
	});
});

describe('buildGoogleFontsFragment', () => {
	it('appends the universal axis spec', () => {
		expect(buildGoogleFontsFragment('ADLaM Display')).toBe(
			'ADLaM Display:ital,wght@0,400;0,700;1,400;1,700',
		);
	});

	it('adds extra weights as sorted upright and italic tuples', () => {
		expect(buildGoogleFontsFragment('DM Sans', [600, 300, 700])).toBe(
			'DM Sans:ital,wght@0,300;0,400;0,600;0,700;1,300;1,400;1,600;1,700',
		);
		expect(buildGoogleFontsFragment('DM Sans', [700, 400])).toBe(
			'DM Sans:ital,wght@0,400;0,700;1,400;1,700',
		);
	});
});

describe('buildGoogleFontsHref', () => {
	it('joins verified fragments with spaces encoded as %20', () => {
		const href = buildGoogleFontsHref([
			buildGoogleFontsFragment('ADLaM Display'),
			buildGoogleFontsFragment('Roboto'),
		]);
		expect(href).toBe(
			'https://fonts.googleapis.com/css2?family=ADLaM%20Display%3Aital%2Cwght%400%2C400%3B0%2C700%3B1%2C400%3B1%2C700&family=Roboto%3Aital%2Cwght%400%2C400%3B0%2C700%3B1%2C400%3B1%2C700&display=swap',
		);
	});

	it('returns null when nothing was verified', () => {
		expect(buildGoogleFontsHref([])).toBeNull();
	});

	it('accepts a self-hosted base URL override', () => {
		const href = buildGoogleFontsHref(
			[buildGoogleFontsFragment('Carlito')],
			'https://fonts.example.com/css2',
		);
		expect(href).toMatch(/^https:\/\/fonts\.example\.com\/css2\?family=Carlito/u);
	});
});

describe('google-fonts-catalogue', () => {
	it('is a sorted, deduplicated snapshot with a regeneration date', () => {
		expect(GOOGLE_FONTS_CATALOGUE_DATE).toMatch(/^\d{4}-\d{2}-\d{2}$/u);
		expect(GOOGLE_FONTS_FAMILIES.length).toBeGreaterThan(1000);
		expect(new Set(GOOGLE_FONTS_FAMILIES).size).toBe(GOOGLE_FONTS_FAMILIES.length);
		expect([...GOOGLE_FONTS_FAMILIES]).toStrictEqual(
			[...GOOGLE_FONTS_FAMILIES].sort((a, b) => a.localeCompare(b, 'en')),
		);
	});
});

describe('findGoogleFontsFamily', () => {
	it('answers with the canonical spelling, case- and whitespace-insensitively', () => {
		expect(findGoogleFontsFamily('ADLaM Display')).toBe('ADLaM Display');
		expect(findGoogleFontsFamily('adlam  display ')).toBe('ADLaM Display');
		expect(findGoogleFontsFamily('roboto')).toBe('Roboto');
	});

	it('rejects families the API does not serve without any request', () => {
		expect(findGoogleFontsFamily('Calibri')).toBeNull();
		expect(findGoogleFontsFamily('Helvetica Neue Medium')).toBeNull();
		expect(findGoogleFontsFamily('')).toBeNull();
	});
});

describe('findMetricCompatibleGoogleFontsFamily', () => {
	it('resolves an unserved Office font to its metric-compatible clone', () => {
		expect(findMetricCompatibleGoogleFontsFamily('Calibri')).toBe('Carlito');
		expect(findMetricCompatibleGoogleFontsFamily('Cambria')).toBe('Caladea');
		expect(findMetricCompatibleGoogleFontsFamily('Arial')).toBe('Arimo');
		expect(findMetricCompatibleGoogleFontsFamily('Times New Roman')).toBe('Tinos');
		expect(findMetricCompatibleGoogleFontsFamily('Courier New')).toBe('Cousine');
		expect(findMetricCompatibleGoogleFontsFamily('Georgia')).toBe('Gelasio');
	});

	it('chains through Aptos to the same Calibri clone', () => {
		expect(findMetricCompatibleGoogleFontsFamily('Aptos')).toBe('Carlito');
	});

	it('returns null when the family is already servable under its own name', () => {
		expect(findMetricCompatibleGoogleFontsFamily('Roboto')).toBeNull();
	});

	it('returns null when no fallback in the chain is on the catalogue', () => {
		// Segoe UI's metric-compatible match (Selawik) is not on Google Fonts.
		expect(findMetricCompatibleGoogleFontsFamily('Segoe UI')).toBeNull();
		expect(findMetricCompatibleGoogleFontsFamily('Totally Unknown Face')).toBeNull();
	});
});

describe('matchGoogleWebfontFragments', () => {
	it('keeps catalogue families (canonically spelled) and drops unknown ones', () => {
		const fragments = matchGoogleWebfontFragments(['adlam display', 'Roboto', 'Totally Unknown']);
		expect(fragments).toStrictEqual([
			'ADLaM Display:ital,wght@0,400;0,700;1,400;1,700',
			'Roboto:ital,wght@0,400;0,700;1,400;1,700',
		]);
	});

	it('never touches the network', () => {
		const fetchSpy = vi.fn();
		vi.stubGlobal('fetch', fetchSpy);
		matchGoogleWebfontFragments(['ADLaM Display', 'Totally Unknown']);
		expect(fetchSpy).not.toHaveBeenCalled();
	});

	it('substitutes an unmatched Office font with its metric-compatible clone', () => {
		expect(matchGoogleWebfontFragments(['Calibri'])).toStrictEqual([
			buildGoogleFontsFragment('Carlito'),
		]);
	});

	it('requests a weight variant from its base family, with the weights the deck uses', () => {
		expect(
			matchGoogleWebfontFragments(['DM Sans Light', 'DM Sans', 'DM Sans Medium', 'DM Sans Black']),
		).toStrictEqual([
			'DM Sans:ital,wght@0,300;0,400;0,500;0,700;0,900;1,300;1,400;1,500;1,700;1,900',
		]);
	});

	it('keeps the universal axis for catalogue families named with a weight word', () => {
		expect(matchGoogleWebfontFragments(['Archivo Black'])).toStrictEqual([
			'Archivo Black:ital,wght@0,400;0,700;1,400;1,700',
		]);
	});

	it('dedupes two families that resolve to the same clone into one fragment', () => {
		// Aptos and Calibri both resolve to Carlito; the request should not
		// ask Google Fonts for the same family twice.
		expect(matchGoogleWebfontFragments(['Aptos', 'Calibri'])).toStrictEqual([
			buildGoogleFontsFragment('Carlito'),
		]);
	});
});

describe('resolveGoogleFontsCss2Base', () => {
	it('keeps the Google Fonts CSS2 API when the option is unset', () => {
		expect(resolveGoogleFontsCss2Base(undefined)).toBe(GOOGLE_FONTS_CSS2_BASE);
	});

	it('turns the fallback off for false', () => {
		expect(resolveGoogleFontsCss2Base(false)).toBeNull();
	});

	it('uses a valid self-hosted mirror base URL', () => {
		expect(resolveGoogleFontsCss2Base({ baseUrl: 'https://fonts.example.test/css2' })).toBe(
			'https://fonts.example.test/css2',
		);
		expect(resolveGoogleFontsCss2Base({ baseUrl: ' http://127.0.0.1:8080/css2 ' })).toBe(
			'http://127.0.0.1:8080/css2',
		);
	});

	it('fails closed with one warning for an invalid base URL', () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		expect(resolveGoogleFontsCss2Base({ baseUrl: 'ftp://fonts.example.test/css2' })).toBeNull();
		expect(resolveGoogleFontsCss2Base({ baseUrl: 'ftp://fonts.example.test/css2' })).toBeNull();
		expect(warn).toHaveBeenCalledTimes(1);
	});
});

describe('normalizeGoogleFontsBaseUrl', () => {
	it.each([
		'',
		'/css2',
		'fonts.example.test/css2',
		'javascript:alert(1)',
		'data:text/css,body{}',
		'https://user:secret@fonts.example.test/css2',
		'https://user@fonts.example.test/css2',
		'https://fonts.example.test/css2?key=1',
		'https://fonts.example.test/css2#x',
	])('rejects %j', (value) => {
		expect(normalizeGoogleFontsBaseUrl(value)).toBeNull();
	});

	it('rejects non-string values', () => {
		expect(normalizeGoogleFontsBaseUrl(undefined)).toBeNull();
		expect(normalizeGoogleFontsBaseUrl(42)).toBeNull();
	});
});

describe('resolveGoogleWebfontHref', () => {
	it('builds the href on a self-hosted mirror with the same query format', async () => {
		const base = resolveGoogleFontsCss2Base({ baseUrl: 'https://fonts.example.test/css2' });
		const href = await resolveGoogleWebfontHref(
			[slide(textEl('ADLaM Display'))],
			[],
			() => false,
			base,
		);
		expect(href).toBe(
			`https://fonts.example.test/css2?family=${encodeURIComponent('ADLaM Display:ital,wght@0,400;0,700;1,400;1,700')}&display=swap`,
		);
	});

	it('resolves nothing and inspects no family when the fallback is off', async () => {
		const isInstalled = vi.fn(() => false);
		await expect(
			resolveGoogleWebfontHref(
				[slide(textEl('ADLaM Display'))],
				[],
				isInstalled,
				resolveGoogleFontsCss2Base(false),
			),
		).resolves.toBeNull();
		expect(isInstalled).not.toHaveBeenCalled();
	});

	it('resolves the href for a deck referencing an unembedded served font', async () => {
		const href = await resolveGoogleWebfontHref([slide(textEl('ADLaM Display'))], []);
		expect(href).toContain('family=ADLaM%20Display%3Aital');
	});

	it('returns null when the family is embedded or truly unmatched', async () => {
		await expect(
			resolveGoogleWebfontHref([slide(textEl('ADLaM Display'))], [{ name: 'ADLaM Display' }]),
		).resolves.toBeNull();
		await expect(
			resolveGoogleWebfontHref([slide(textEl('Totally Unknown Face'))], []),
		).resolves.toBeNull();
	});

	it('requests an embedded family whose files do not match it', async () => {
		// The embedded entry has no readable file, so it is trusted; see
		// `embedded-font-trust.test.ts` for the files that are not.
		await expect(
			resolveGoogleWebfontHref([slide(textEl('DM Sans Light'))], [{ name: 'DM Sans Light' }]),
		).resolves.toBeNull();
		const href = await resolveGoogleWebfontHref([slide(textEl('DM Sans Light'))], [], () => false);
		expect(href).toContain(`family=${encodeURIComponent('DM Sans:ital,wght@0,300;0,400;0,700')}`);
	});

	it('resolves an unembedded Office font to its metric-compatible clone', async () => {
		const href = await resolveGoogleWebfontHref([slide(textEl('Calibri'))], []);
		expect(href).toContain('family=Carlito%3Aital');
	});

	it('an embedded Office font is NOT redirected to its clone', async () => {
		await expect(
			resolveGoogleWebfontHref([slide(textEl('Calibri'))], [{ name: 'Calibri' }]),
		).resolves.toBeNull();
	});

	it('accepts a self-hosted CSS2 API mirror base URL', async () => {
		const href = await resolveGoogleWebfontHref(
			[slide(textEl('Calibri'))],
			[],
			undefined,
			'https://fonts.example.com/css2',
		);
		expect(href).toMatch(/^https:\/\/fonts\.example\.com\/css2\?family=Carlito/u);
	});

	it('never loads a family the runtime reports as installed', async () => {
		await expect(
			resolveGoogleWebfontHref(
				[slide(textEl('Roboto'), textEl('ADLaM Display'))],
				[],
				(family) => family === 'Roboto',
			),
		).resolves.toBe(buildGoogleFontsHref([buildGoogleFontsFragment('ADLaM Display')]));
	});

	it('keeps a family it already resolved even when it now measures as installed', async () => {
		// Once the injected stylesheet has loaded, the canvas measurement sees
		// the webfont itself and reports the family as "installed". Without the
		// session guard the second resolve would drop the family, the binding
		// would remove the <link>, and the third resolve would find it missing
		// again: an oscillation that re-fetches on every edit.
		const slides = [slide(textEl('ADLaM Display'))];
		let webfontLoaded = false;
		const isInstalled = (): boolean => webfontLoaded;

		const first = await resolveGoogleWebfontHref(slides, [], isInstalled);
		expect(first).toContain('family=ADLaM%20Display');
		webfontLoaded = true;

		const second = await resolveGoogleWebfontHref(slides, [], isInstalled);
		expect(second).toBe(first);
	});

	it('still runs the local check for families the session has not resolved', async () => {
		await resolveGoogleWebfontHref([slide(textEl('ADLaM Display'))], [], () => false);
		await expect(
			resolveGoogleWebfontHref(
				[slide(textEl('ADLaM Display'), textEl('Roboto'))],
				[],
				(family) => family === 'Roboto',
			),
		).resolves.toBe(buildGoogleFontsHref([buildGoogleFontsFragment('ADLaM Display')]));
	});
});

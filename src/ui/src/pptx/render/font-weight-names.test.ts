import { describe, expect, it } from 'vitest';

import {
	fontStyleShorthand,
	fontWeightAttribute,
	fontWeightShorthand,
	parseFontWeightName,
	resolveFontItalic,
	resolveFontWeight,
	resolveGoogleFontsFace,
} from './font-weight-names';

describe('parseFontWeightName', () => {
	it('splits a trailing weight off the family', () => {
		expect(parseFontWeightName('DM Sans Light')).toEqual({
			family: 'DM Sans',
			weight: 300,
			italic: false,
		});
		expect(parseFontWeightName('Open Sans SemiBold')).toEqual({
			family: 'Open Sans',
			weight: 600,
			italic: false,
		});
		expect(parseFontWeightName('Montserrat ExtraBold')).toEqual({
			family: 'Montserrat',
			weight: 800,
			italic: false,
		});
		expect(parseFontWeightName('Inter Thin')).toEqual({
			family: 'Inter',
			weight: 100,
			italic: false,
		});
		expect(parseFontWeightName('Roboto Black')).toEqual({
			family: 'Roboto',
			weight: 900,
			italic: false,
		});
	});

	it('reads spaced and hyphenated spellings of compound weights', () => {
		expect(parseFontWeightName('Poppins Semi Bold')).toEqual({
			family: 'Poppins',
			weight: 600,
			italic: false,
		});
		expect(parseFontWeightName('Poppins Extra-Light')).toEqual({
			family: 'Poppins',
			weight: 200,
			italic: false,
		});
		expect(parseFontWeightName('Poppins Semi-Bold')).toEqual({
			family: 'Poppins',
			weight: 600,
			italic: false,
		});
	});

	it('reads a trailing style token', () => {
		expect(parseFontWeightName('Lato Light Italic')).toEqual({
			family: 'Lato',
			weight: 300,
			italic: true,
		});
		expect(parseFontWeightName('Lato Italic')).toEqual({
			family: 'Lato',
			weight: 400,
			italic: true,
		});
		expect(parseFontWeightName('Lato Bold Oblique')).toEqual({
			family: 'Lato',
			weight: 700,
			italic: true,
		});
	});

	it('returns null for a name without weight or style tokens', () => {
		expect(parseFontWeightName('DM Sans')).toBeNull();
		expect(parseFontWeightName('Light')).toBeNull();
		expect(parseFontWeightName('')).toBeNull();
	});
});

describe('resolveGoogleFontsFace', () => {
	it('serves a style variant from the italic faces of its base family', () => {
		expect(resolveGoogleFontsFace('Lato Light Italic')).toEqual({
			family: 'Lato',
			weight: 300,
			italic: true,
		});
		expect(resolveGoogleFontsFace('Lato Italic')).toEqual({
			family: 'Lato',
			weight: 400,
			italic: true,
		});
	});

	it('serves a weight variant from its base family', () => {
		expect(resolveGoogleFontsFace('DM Sans Light')).toEqual({ family: 'DM Sans', weight: 300 });
		expect(resolveGoogleFontsFace('dm sans  medium')).toEqual({ family: 'DM Sans', weight: 500 });
	});

	it('prefers a catalogue family named with a weight word', () => {
		expect(resolveGoogleFontsFace('Archivo Black')).toEqual({ family: 'Archivo Black' });
		expect(resolveGoogleFontsFace('DM Sans')).toEqual({ family: 'DM Sans' });
	});

	it('returns null when neither the name nor its base is served', () => {
		expect(resolveGoogleFontsFace('Arial Rounded MT Bold')).toBeNull();
		expect(resolveGoogleFontsFace('Contoso Sans Light')).toBeNull();
	});
});

describe('resolveFontWeight', () => {
	it('renders weight variants of catalogue families at their weight', () => {
		expect(resolveFontWeight('DM Sans Light', false)).toBe(300);
		expect(resolveFontWeight('DM Sans Medium', undefined)).toBe(500);
		expect(resolveFontWeight('DM Sans SemiBold', false)).toBe(600);
	});

	it('keeps bold at least bold', () => {
		expect(resolveFontWeight('DM Sans Light', true)).toBe(700);
		expect(resolveFontWeight('DM Sans Black', true)).toBe(900);
		expect(resolveFontWeight('DM Sans Black', false)).toBe(900);
	});

	it('keeps the plain mapping for every other family', () => {
		expect(resolveFontWeight('DM Sans', false)).toBe(400);
		expect(resolveFontWeight('DM Sans', true)).toBe(700);
		// Single-weight families whose names end in a weight word: asking for
		// another weight would make the browser synthesise one.
		expect(resolveFontWeight('Archivo Black', false)).toBe(400);
		expect(resolveFontWeight('Arial Rounded MT Bold', false)).toBe(400);
		expect(resolveFontWeight('Arial Black', false)).toBe(400);
		expect(resolveFontWeight(undefined, true)).toBe(700);
		expect(resolveFontWeight('', false)).toBe(400);
	});

	it('reads the first entry of a resolved CSS font-family list', () => {
		expect(resolveFontWeight('"DM Sans Light", "DM Sans", sans-serif', false)).toBe(300);
		expect(resolveFontWeight("'DM Sans Medium', sans-serif", false)).toBe(500);
		expect(resolveFontWeight('DM Sans Light, sans-serif', false)).toBe(300);
		expect(resolveFontWeight('"DM Sans", "DM Sans Light"', false)).toBe(400);
	});
});

describe('weight serialisation', () => {
	it('keeps the keywords for regular and bold', () => {
		expect(fontWeightAttribute('DM Sans', false)).toBeUndefined();
		expect(fontWeightAttribute('DM Sans', true)).toBe('bold');
		expect(fontWeightShorthand('DM Sans', false)).toBe('');
		expect(fontWeightShorthand('DM Sans', true)).toBe('bold ');
	});

	it('writes any other weight as a number', () => {
		expect(fontWeightAttribute('DM Sans Light', false)).toBe('300');
		expect(fontWeightShorthand('"DM Sans Light", sans-serif', false)).toBe('300 ');
		expect(fontWeightShorthand('DM Sans Black', true)).toBe('900 ');
	});
});

describe('resolveFontItalic', () => {
	it('renders style variants of catalogue families italic', () => {
		expect(resolveFontItalic('Lato Light Italic', false)).toBe(true);
		expect(resolveFontItalic('Lato Italic', undefined)).toBe(true);
		expect(resolveFontItalic('"Lato Light Italic", "Lato", sans-serif', false)).toBe(true);
		expect(resolveFontWeight('Lato Light Italic', false)).toBe(300);
	});

	it('keeps the run style for every other family', () => {
		expect(resolveFontItalic('Lato Light', false)).toBe(false);
		expect(resolveFontItalic('Lato', true)).toBe(true);
		// Not a style variant of a family Google serves.
		expect(resolveFontItalic('Acme Brand Italic', false)).toBe(false);
		expect(resolveFontItalic(undefined, false)).toBe(false);
	});

	it('writes the style into a font shorthand', () => {
		expect(fontStyleShorthand('Lato Light Italic', false)).toBe('italic ');
		expect(fontStyleShorthand('Lato', true)).toBe('italic ');
		expect(fontStyleShorthand('Lato', false)).toBe('');
	});
});

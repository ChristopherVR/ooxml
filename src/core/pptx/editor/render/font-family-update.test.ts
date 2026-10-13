import { describe, expect, it } from 'vitest';
import { fontFamilyStyleUpdates, mergeTextStyleUpdate } from './font-family-update';
const sameFace = {
	fontFamily: 'Old Sans',
	eastAsiaFont: 'Old Sans',
	complexScriptFont: 'Old Sans',
};

describe('fontFamilyStyleUpdates', () => {
	it('carries the east Asian and complex script faces that matched the Latin face', () => {
		expect(fontFamilyStyleUpdates(sameFace, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
			complexScriptFont: 'New Serif',
		});
	});

	it('keeps a distinct east Asian face the run chose on purpose', () => {
		const style = { fontFamily: 'Latin Sans', eastAsiaFont: 'Asian Gothic' };
		expect(mergeTextStyleUpdate(style, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'Asian Gothic',
		});
	});

	it('follows a theme face that resolved to the Latin face and drops the replaced tokens', () => {
		const style = {
			fontFamily: 'Theme Body',
			latinFontThemeToken: '+mn-lt',
			eastAsiaFont: 'Theme Body',
			eastAsiaFontThemeToken: '+mn-ea',
		};
		expect(mergeTextStyleUpdate(style, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			latinFontThemeToken: undefined,
			eastAsiaFont: 'New Serif',
			eastAsiaFontThemeToken: undefined,
		});
	});

	it('keeps a theme east Asian face that differs from the theme Latin face', () => {
		const style = {
			fontFamily: 'Theme Latin',
			latinFontThemeToken: '+mn-lt',
			eastAsiaFont: 'Theme Gothic',
			eastAsiaFontThemeToken: '+mn-ea',
		};
		expect(mergeTextStyleUpdate(style, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			latinFontThemeToken: undefined,
			eastAsiaFont: 'Theme Gothic',
			eastAsiaFontThemeToken: '+mn-ea',
		});
	});

	it('adds no script faces to a run that only names a Latin face', () => {
		expect(
			fontFamilyStyleUpdates({ fontFamily: 'Old Sans' }, { fontFamily: 'New Serif' }),
		).toStrictEqual({
			fontFamily: 'New Serif',
		});
		expect(fontFamilyStyleUpdates(undefined, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
		});
	});

	it('follows a script token that matches the Latin token without a concrete face', () => {
		const style = {
			fontFamily: 'Theme Body',
			latinFontThemeToken: '+mn-lt',
			eastAsiaFontThemeToken: '+mn-lt',
			bold: true,
		};
		expect(mergeTextStyleUpdate(style, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			latinFontThemeToken: undefined,
			eastAsiaFont: 'New Serif',
			eastAsiaFontThemeToken: undefined,
			bold: true,
		});
	});

	it('compares an inherited Latin face against the text box style', () => {
		const run = { eastAsiaFont: 'Old Sans', italic: true };
		expect(mergeTextStyleUpdate(run, { fontFamily: 'New Serif' }, sameFace)).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
			italic: true,
		});
		const distinct = { eastAsiaFont: 'Asian Gothic' };
		expect(mergeTextStyleUpdate(distinct, { fontFamily: 'New Serif' }, sameFace)).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'Asian Gothic',
		});
	});

	it('prefers the run Latin face over the inherited one', () => {
		const run = { fontFamily: 'Run Latin', eastAsiaFont: 'Old Sans' };
		expect(mergeTextStyleUpdate(run, { fontFamily: 'New Serif' }, sameFace)).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'Old Sans',
		});
	});

	it('matches an inherited Latin theme token', () => {
		const box = { fontFamily: 'Theme Body', latinFontThemeToken: '+mn-lt' };
		const run = { eastAsiaFontThemeToken: '+mn-lt' };
		expect(mergeTextStyleUpdate(run, { fontFamily: 'New Serif' }, box)).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
			eastAsiaFontThemeToken: undefined,
		});
	});

	it('leaves other updates and explicit script faces alone', () => {
		const bold = { bold: true };
		expect(fontFamilyStyleUpdates(sameFace, bold)).toBe(bold);
		expect(
			fontFamilyStyleUpdates(sameFace, { fontFamily: 'New Serif', eastAsiaFont: 'Picked EA' }),
		).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'Picked EA',
			complexScriptFont: 'New Serif',
		});
	});
});

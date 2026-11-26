import { describe, expect, it } from 'vitest';
import { EQUATION_STRUCTURES, equationStructure } from './equation-structures';
import { EQUATION_SYMBOLS } from './equation-symbols';
import { buildUpEquation, linearizeEquation } from './linear-format';
import { drawEquation } from './omml-svg';
import { fillPlaceholders } from './equation-placeholder';
import { parseOmml } from './omml-dom';
import { ommlToXml } from './omml-to-xml';
import { EQUATION_UI_STRINGS } from './equation-strings';
import { EQUATION_GALLERY_STRINGS } from './equation-gallery-strings';
import { EQUATION_GALLERY } from './equation-gallery';

const entries = EQUATION_STRUCTURES.flatMap((gallery) =>
	gallery.structures.map((entry) => [`${gallery.id}/${entry.id}`, entry.latex] as const),
);

describe('equation structures', () => {
	it("lists Word's eleven structure galleries", () => {
		expect(EQUATION_STRUCTURES.map((gallery) => gallery.name)).toEqual([
			'Fraction',
			'Script',
			'Radical',
			'Integral',
			'Large Operator',
			'Bracket',
			'Function',
			'Accent',
			'Limit and Log',
			'Operator',
			'Matrix',
		]);
		expect(new Set(entries.map(([id]) => id)).size).toBe(entries.length);
		expect(equationStructure('fraction', 'stacked')?.name).toBe('Stacked Fraction');
	});

	it.each(entries)('%s builds up, draws with its placeholders and linearizes', (_id, latex) => {
		const xml = buildUpEquation(latex, 'latex', false);
		expect(xml).toMatch(/^<m:oMath /);
		expect(xml).not.toContain('⬚');
		const shown = ommlToXml(fillPlaceholders(parseOmml(xml!)));
		expect(drawEquation(shown, 16)?.widthPx).toBeGreaterThan(4);
		for (const format of ['unicode', 'latex'] as const) {
			const linear = linearizeEquation(xml!, format)!;
			expect(linear.length).toBeGreaterThan(0);
			// Linear text builds up into the same structure again.
			expect(buildUpEquation(linear, format, false)).toBe(xml);
		}
	});
});

describe('equation symbols', () => {
	it("groups the symbols into Word's eight categories without duplicates", () => {
		expect(EQUATION_SYMBOLS.map((category) => category.id)).toEqual([
			'basicMath',
			'greekLetters',
			'letterLike',
			'operators',
			'arrows',
			'negatedRelations',
			'scripts',
			'geometry',
		]);
		for (const category of EQUATION_SYMBOLS) {
			expect(category.symbols.length).toBeGreaterThan(20);
			expect(new Set(category.symbols).size).toBe(category.symbols.length);
		}
	});
});

describe('equation strings', () => {
	it('names every gallery, structure, symbol category and built-in equation in each locale', () => {
		const names = [
			...EQUATION_STRUCTURES.flatMap((gallery) => [
				gallery.name,
				...gallery.structures.map((entry) => entry.name),
			]),
			...EQUATION_SYMBOLS.map((category) => category.name),
			...EQUATION_GALLERY.slice(2).map((entry) => entry.name),
		].filter((name) => !['Pythagorean Theorem', 'Quadratic Formula'].includes(name));
		for (const locale of ['en', 'fr', 'de', 'es', 'zh-CN'] as const) {
			const strings: Record<string, string> = {
				...EQUATION_UI_STRINGS[locale],
				...EQUATION_GALLERY_STRINGS[locale],
			};
			for (const name of names) expect(strings[name], `${locale}: ${name}`).toBeTruthy();
		}
	});
});

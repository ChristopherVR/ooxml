import { describe, expect, it } from 'vitest';
import { convertOmmlToLatex, parseOmml } from './index';
import { EQUATION_GALLERY } from './equation-gallery';
import { latexToOmmlXml } from './omml-to-xml';
import { drawEquation } from './omml-svg';

describe('equation gallery', () => {
	it('lists the four built-in equations', () => {
		expect(EQUATION_GALLERY.map((entry) => entry.id)).toEqual([
			'area',
			'quadratic',
			'pythagorean',
			'binomial',
		]);
	});

	it.each(EQUATION_GALLERY.map((entry) => [entry.name, entry.latex] as const))(
		'%s converts to Office Math, reads back and draws',
		(_name, latex) => {
			for (const display of [false, true]) {
				const xml = latexToOmmlXml(latex, display)!;
				expect(xml).toContain(display ? '<m:oMathPara' : '<m:oMath');
				expect(convertOmmlToLatex(parseOmml(xml)).length).toBeGreaterThan(3);
				const drawing = drawEquation(xml, 16)!;
				expect(drawing.widthPx).toBeGreaterThan(30);
				expect(drawing.svg).toContain('<text');
			}
		},
	);

	it('the binomial theorem keeps its sum, its binomial coefficient and both scripts', () => {
		const xml = latexToOmmlXml(EQUATION_GALLERY[3]!.latex, true)!;
		expect(xml).toContain('<m:nary>');
		expect(xml).toContain('<m:f>');
		expect(xml).toContain('<m:sSup>');
		expect(xml).toContain('m:val="noBar"');
	});
});

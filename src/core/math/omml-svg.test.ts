import { describe, expect, it } from 'vitest';
import { convertOmmlToLatex, convertOmmlToMathMl, parseOmml } from './index';
import { latexToOmmlXml } from './omml-to-xml';
import { drawEquation } from './omml-svg';

const draw = (latex: string, size = 16, display = false) => {
	const xml = latexToOmmlXml(latex, display);
	if (!xml) throw new Error('no markup');
	return drawEquation(xml, size, latex);
};

describe('latexToOmmlXml', () => {
	it('writes inline and display markup that reads back to the same LaTeX', () => {
		for (const latex of [
			'x^2+y^2=z^2',
			'\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}',
			'\\sum_{i=1}^{n} i',
			'\\int_0^1 x\\,dx',
			'\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}',
		]) {
			const inline = latexToOmmlXml(latex, false)!;
			expect(inline.startsWith('<m:oMath xmlns:m=')).toBe(true);
			const display = latexToOmmlXml(latex, true)!;
			expect(display.startsWith('<m:oMathPara xmlns:m=')).toBe(true);
			expect(convertOmmlToLatex(parseOmml(inline))).toBe(convertOmmlToLatex(parseOmml(display)));
			// The LaTeX read back is stable: writing it again gives the same equation.
			const back = convertOmmlToLatex(parseOmml(inline));
			expect(convertOmmlToLatex(parseOmml(latexToOmmlXml(back, false)!))).toBe(back);
		}
		expect(latexToOmmlXml('   ', false)).toBeUndefined();
	});

	it('escapes markup characters in text', () => {
		const xml = latexToOmmlXml('a<b', false)!;
		expect(xml).toContain('&lt;');
		expect(() => parseOmml(xml)).not.toThrow();
	});
});

describe('equation drawing', () => {
	it('sizes an equation from its box model and reports the descent', () => {
		const letter = draw('x')!;
		expect(letter.widthPx).toBeGreaterThan(8);
		expect(letter.heightPx).toBeGreaterThan(16);
		expect(letter.svg).toContain('<svg');
		expect(letter.svg).toContain('aria-label="x"');
		const bigger = draw('x', 32)!;
		expect(bigger.widthPx).toBeGreaterThan(letter.widthPx * 1.8);
	});

	it('stacks fractions, raises scripts and widens with content', () => {
		const flat = draw('a+b')!;
		const fraction = draw('\\frac{a}{b}')!;
		expect(fraction.heightPx).toBeGreaterThan(flat.heightPx * 1.6);
		expect(fraction.svg).toContain('<line');
		expect(fraction.descentPx).toBeGreaterThan(flat.descentPx);
		const power = draw('x^2')!;
		expect(power.heightPx).toBeGreaterThan(draw('x')!.heightPx);
		expect(draw('a+b+c+d')!.widthPx).toBeGreaterThan(flat.widthPx);
	});

	it('draws radicals, n-ary operators, delimiters and matrices', () => {
		expect(draw('\\sqrt{x}')!.svg).toContain('<path');
		const sum = draw('\\sum_{i=1}^{n} i')!;
		expect(sum.svg).toContain('∑');
		expect(sum.heightPx).toBeGreaterThan(draw('i')!.heightPx * 2);
		const tall = draw('\\left(\\frac{a}{b}\\right)')!;
		expect(tall.svg).toContain('scale(1 ');
		const matrix = draw('\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}')!;
		expect(matrix.heightPx).toBeGreaterThan(draw('a')!.heightPx * 1.8);
	});

	it('returns undefined for markup without content or that cannot be read', () => {
		expect(drawEquation('<m:oMath xmlns:m="x"/>', 16)).toBeUndefined();
		expect(drawEquation('not xml', 16)).toBeUndefined();
	});
});

describe('minus signs', () => {
	const widthOf = (latex: string) => draw(latex)!.widthPx;

	it('draws a hyphen-minus as the minus sign', () => {
		const svg = draw('a-b')!.svg;
		expect(svg).toContain('−');
		expect(svg).not.toMatch(/>[^<]*-[^<]*</);
	});

	it('spaces a binary minus but not a unary one or one in a script', () => {
		expect(widthOf('a-b')).toBeCloseTo(widthOf('a+b'), 0);
		// "-b" has no operand to its left, so it is narrower than "a-b" by a letter and both gaps.
		expect(widthOf('a-b') - widthOf('-b')).toBeGreaterThan(16 * 0.5 + 2 * 0.22 * 16 - 1);
		// After a relation the sign is unary too.
		expect(widthOf('x=-b')).toBeLessThan(widthOf('x=a-b'));
		// In a script no operator spacing is added.
		const minus = widthOf('x^{n-k}') - widthOf('x');
		const plus = widthOf('x^{n+k}') - widthOf('x');
		expect(minus).toBeCloseTo(plus, 0);
		expect(minus).toBeLessThan(16 * 0.7 * 2.6);
	});

	it('shows the minus sign in the MathML preview too', () => {
		const mathml = convertOmmlToMathMl(parseOmml(latexToOmmlXml('a-b', false)!));
		expect(mathml).toContain('<mo>−</mo>');
	});
});

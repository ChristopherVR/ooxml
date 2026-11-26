import { describe, expect, it } from 'vitest';
import {
	buildUpEquation,
	emptyEquationXml,
	linearDisplayText,
	linearEquationXml,
	linearizeEquation,
	mathAutoCorrect,
} from './linear-format';
import { unicodeMathToLatex } from './unicode-math';

const strip = (xml: string | undefined) => xml?.replace(/ xmlns:m="[^"]*"/, '');
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

describe('unicodeMathToLatex', () => {
	it.each([
		['a/b', String.raw`\frac{a}{b}`],
		['(a+b)/c', String.raw`\frac{a+b}{c}`],
		['a+b/c', String.raw`a+\frac{b}{c}`],
		['x^2', 'x^{2}'],
		['2x^2', '2x^{2}'],
		['10^3', '{10}^{3}'],
		['x_i^2', 'x_{i}^{2}'],
		['e^-x', 'e^{-x}'],
		['√x', String.raw`\sqrt{x}`],
		['√(3&x)', String.raw`\sqrt[3]{x}`],
		['∛x', String.raw`\sqrt[3]{x}`],
		['∑_(i=1)^n▒a_i', String.raw`\sum_{i=1}^{n}{a_{i}}`],
		['\\alpha+\\beta', 'α+β'],
		['sin x', String.raw`\sin{x}`],
		['lim_(n→∞) a_n', String.raw`\lim_{n→∞}{a_{n}}`],
		['|x|', String.raw`\left|x\right|`],
		['■(a&b@c&d)', String.raw`\begin{matrix}a&b\\c&d\end{matrix}`],
		['"if "x', String.raw`\text{if }x`],
		['a¦b', String.raw`{a \atop b}`],
		['⬚/⬚', String.raw`\frac{⬚}{⬚}`],
		['3.14/2', String.raw`\frac{3.14}{2}`],
		['{a}', String.raw`\left\{a\right\}`],
		['a&b', String.raw`a\&b`],
	])('builds up %s', (linear, latex) => {
		expect(unicodeMathToLatex(linear)).toBe(latex);
	});

	it('never throws on unbalanced or stray input', () => {
		for (const text of ['(', ')', 'a/', '^', '√', '∑', '■(', '"', '\\', '(a]', '├a┤', '//'])
			expect(() => unicodeMathToLatex(text)).not.toThrow();
	});
});

describe('buildUpEquation', () => {
	it('builds Unicode linear text into Office Math structures', () => {
		expect(strip(buildUpEquation('a/b', 'unicode', false))).toBe(
			'<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>',
		);
	});

	it('wraps a display equation in m:oMathPara', () => {
		expect(buildUpEquation('x', 'unicode', true)).toMatch(/^<m:oMathPara [^>]*><m:oMath>/);
	});

	it('saves placeholders as empty arguments', () => {
		expect(strip(buildUpEquation('⬚/⬚', 'unicode', false))).toBe(
			'<m:oMath><m:f><m:num/><m:den/></m:f></m:oMath>',
		);
		expect(strip(buildUpEquation(String.raw`\sqrt{⬚}`, 'latex', false))).toBe(
			'<m:oMath><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e/></m:rad></m:oMath>',
		);
	});

	it('writes properties before content, as the schema orders them', () => {
		const xml = buildUpEquation(String.raw`\text{if}\binom{n}{k}`, 'latex', false)!;
		expect(xml).toContain('<m:r><m:rPr><m:nor m:val="1"/></m:rPr><m:t>if</m:t></m:r>');
		expect(xml).toContain('<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>');
	});

	it('reports LaTeX it cannot convert and builds empty text as an empty equation', () => {
		expect(buildUpEquation('x'.repeat(20_000), 'latex', false)).toBeUndefined();
		expect(buildUpEquation('  ', 'unicode', true)).toBe(emptyEquationXml(true));
	});
});

describe('linearizeEquation', () => {
	it.each([
		'a/b',
		'(a+b)/c',
		'x^2',
		'x_i^2',
		'√x',
		'√(3&x)',
		'∑_(i=1)^n▒a_i',
		'sin x',
		'|x|',
		'■(a&b@c&d)',
		'"if "x',
		'⬚/⬚',
		'a¦b',
		'x┬n',
		'a∕b',
		'_a^b x',
		'=┴"def"',
		'sin^(-1)⁡x',
		'log_2⁡x',
	])('round-trips %s through Office Math', (linear) => {
		const xml = buildUpEquation(linear, 'unicode', false)!;
		const back = linearizeEquation(xml, 'unicode')!;
		expect(buildUpEquation(back, 'unicode', false)).toBe(xml);
	});

	it('shows empty arguments as placeholders in both formats', () => {
		const xml = buildUpEquation('⬚/⬚', 'unicode', false)!;
		expect(linearizeEquation(xml, 'unicode')).toBe('⬚/⬚');
		expect(linearizeEquation(xml, 'latex')).toBe(String.raw`\frac{⬚}{⬚}`);
	});
});

describe('Linear display', () => {
	it('saves linear text with each operator in its own run, as Word does, and recognises it', () => {
		const xml = linearEquationXml('a/b+x^2', false);
		expect(strip(xml)).toBe(
			'<m:oMath><m:r><m:t>a</m:t></m:r><m:r><m:t>/</m:t></m:r><m:r><m:t>b+x</m:t></m:r><m:r><m:t>^</m:t></m:r><m:r><m:t>2</m:t></m:r></m:oMath>',
		);
		expect(linearDisplayText(xml)).toBe('a/b+x^2');
		expect(linearDisplayText(buildUpEquation('a/b', 'unicode', false)!)).toBeUndefined();
		expect(linearDisplayText(linearEquationXml('x+1', false))).toBeUndefined();
	});

	it('keeps a Professional run holding an operator character (km/h) as text', () => {
		const xml = `<m:oMath xmlns:m="${M}"><m:r><m:t>km/h</m:t></m:r></m:oMath>`;
		expect(linearDisplayText(xml)).toBeUndefined();
		const linear = linearizeEquation(xml, 'unicode')!;
		expect(linear).toBe('km\\/h');
		expect(buildUpEquation(linear, 'unicode', false)).not.toContain('<m:f>');
	});
});

describe('linear format details', () => {
	it('nests repeated scripts to the right, as Word builds them up', () => {
		expect(unicodeMathToLatex('x^2^3')).toBe('x^{2^{3}}');
		expect(unicodeMathToLatex('a_b_c')).toBe('a_{b_{c}}');
		expect(buildUpEquation('x^2^3', 'unicode', false)).not.toContain('m:sPre');
	});

	it('reads escaped quotes inside normal text and │ as a bracket separator', () => {
		expect(unicodeMathToLatex('"say \\"hi\\""')).toBe(String.raw`\text{say "hi"}`);
		const xml = buildUpEquation('(a│b)', 'unicode', false)!;
		expect(xml.match(/<m:e>/g)).toHaveLength(2);
		expect(linearizeEquation(xml, 'unicode')).toBe('(a│b)');
	});

	it('returns undefined for markup it cannot read', () => {
		expect(linearizeEquation('<broken', 'unicode')).toBeUndefined();
	});
});

describe('mathAutoCorrect', () => {
	it('replaces the control word before the caret', () => {
		expect(mathAutoCorrect('x=\\alpha', 8)).toEqual({ text: 'x=α', caret: 3 });
		expect(mathAutoCorrect('\\sum', 4)).toEqual({ text: '∑', caret: 1 });
		expect(mathAutoCorrect('\\nothing', 8)).toBeUndefined();
		expect(mathAutoCorrect('abc', 3)).toBeUndefined();
	});
});

import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { convertLatexToOmml } from './latex-to-omml';
import { MAX_LATEX_LENGTH, MAX_OMML_LENGTH } from './limits';
import { ommlFromElement, parseOmml } from './omml-dom';

describe('LaTeX nesting limit', () => {
	it('rejects 1,000 opening braces with a clear error instead of overflowing the stack', () => {
		expect(() => convertLatexToOmml('{'.repeat(1_000))).toThrow(/LaTeX nesting exceeds 512/);
	});

	it('rejects a 1,000-deep superscript chain', () => {
		expect(() => convertLatexToOmml('x' + '^'.repeat(1_000))).toThrow(/nesting exceeds/);
	});

	it('still converts ordinary nested groups and long whitespace runs', () => {
		const nested = `${'{'.repeat(100)}x${'}'.repeat(100)}`;
		expect(convertLatexToOmml(nested)).toHaveProperty('m:oMathPara');
		expect(convertLatexToOmml(`${' '.repeat(1_000)}x`)).toHaveProperty('m:oMathPara');
	});
	it('rejects oversized equations before tokenizing and rejects stray closing braces', () => {
		expect(() => convertLatexToOmml('x'.repeat(MAX_LATEX_LENGTH + 1))).toThrow(/size limit/);
		expect(() => convertLatexToOmml('x}')).toThrow(/closing brace/);
		expect(() => convertLatexToOmml('\\sqrt[}]')).toThrow(/closing brace/);
	});
});

describe('OMML limits', () => {
	it('rejects oversized markup before parsing', () => {
		expect(() => parseOmml('x'.repeat(MAX_OMML_LENGTH + 1))).toThrow(/size limit/);
	});
	it('rejects an element tree deeper than the converter limit', () => {
		const depth = 600;
		const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
		const xml = `<m:oMath xmlns:m="${ns}">${'<m:r>'.repeat(depth)}${'</m:r>'.repeat(depth)}</m:oMath>`;
		const element = parseXml(xml, { limits: { maxDepth: 1_000 } }).documentElement;
		expect(() => ommlFromElement(element)).toThrow(/OMML nesting exceeds 512/);
	});
});

describe('linear format nesting limit', () => {
	it('rejects 16,000 nested brackets with a clear error instead of overflowing the stack', async () => {
		const { unicodeMathToLatex, MAX_LINEAR_NESTING } = await import('./unicode-math');
		for (const open of ['(', '〖', '√', '∑'])
			expect(() => unicodeMathToLatex(open.repeat(16_000))).toThrow(
				`Linear format nesting exceeds ${MAX_LINEAR_NESTING} levels`,
			);
	});
});

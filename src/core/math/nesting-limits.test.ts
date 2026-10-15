import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { convertLatexToOmml } from './latex-to-omml';
import { ommlFromElement } from './omml-dom';

describe('LaTeX nesting limit', () => {
	it('rejects 100,000 opening braces with a clear error instead of overflowing the stack', () => {
		expect(() => convertLatexToOmml('{'.repeat(100_000))).toThrow(/LaTeX nesting exceeds 512/);
	});

	it('rejects a 100,000-deep superscript chain', () => {
		expect(() => convertLatexToOmml('x' + '^'.repeat(100_000))).toThrow(/nesting exceeds/);
	});

	it('still converts ordinary nested groups and long whitespace runs', () => {
		const nested = `${'{'.repeat(100)}x${'}'.repeat(100)}`;
		expect(convertLatexToOmml(nested)).toHaveProperty('m:oMathPara');
		expect(convertLatexToOmml(`${' '.repeat(100_000)}x`)).toHaveProperty('m:oMathPara');
	});
});

describe('OMML DOM depth guard', () => {
	it('rejects an element tree deeper than the converter limit', () => {
		const depth = 600;
		const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
		const xml = `<m:oMath xmlns:m="${ns}">${'<m:r>'.repeat(depth)}${'</m:r>'.repeat(depth)}</m:oMath>`;
		const element = parseXml(xml, { limits: { maxDepth: 1_000 } }).documentElement;
		expect(() => ommlFromElement(element)).toThrow(/OMML nesting exceeds 512/);
	});
});

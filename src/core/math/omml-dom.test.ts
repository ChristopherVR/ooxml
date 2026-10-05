import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import {
	convertLatexToOmml,
	convertOmmlToLatex,
	convertOmmlToMathMl,
	ommlFromElement,
	parseOmml,
} from './index.js';

const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
describe('format-neutral OMML DOM input', () => {
	it('keeps interleaved equation order with arbitrary namespace prefixes', () => {
		const source = `<q:oMath xmlns:q="${ns}"><q:sSup><q:e><q:r><q:t>a</q:t></q:r></q:e><q:sup><q:r><q:t>2</q:t></q:r></q:sup></q:sSup><q:r><q:t>+</q:t></q:r><q:sSup><q:e><q:r><q:t>b</q:t></q:r></q:e><q:sup><q:r><q:t>2</q:t></q:r></q:sup></q:sSup></q:oMath>`;
		const dom = parseXml(source).documentElement;
		const adapted = ommlFromElement(dom);
		const generated = convertLatexToOmml('a^2+b^2');
		expect(convertOmmlToLatex(adapted)).toBe('a^{2}+b^{2}');
		expect(convertOmmlToMathMl(adapted)).toBe(convertOmmlToMathMl(generated));
	});
	it('keeps consecutive runs, normal-text properties and escaped text', () => {
		const adapted = parseOmml(
			`<m:oMath xmlns:m="${ns}"><m:r><m:rPr><m:nor m:val="1"/></m:rPr><m:t>A &amp; B</m:t></m:r><m:r><m:t>&lt;x&gt;</m:t></m:r></m:oMath>`,
		);
		expect(convertOmmlToMathMl(adapted)).toContain('<mi mathvariant="normal">A &amp; B</mi>');
		expect(convertOmmlToMathMl(adapted)).toContain('&lt;x&gt;');
	});
	it('rejects malformed XML and DTD input using the shared parser', () => {
		expect(() => parseOmml('<m:oMath>')).toThrow();
		expect(() => parseOmml('<!DOCTYPE x><x/>')).toThrow(/DTD/);
	});
});

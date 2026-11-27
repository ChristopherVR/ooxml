import { describe, expect, it } from 'vitest';
import { equationParts, equationReadOnlyReason, replaceEquationPart } from './equation-parts';
import { editableEquation, linearizeEquation, writeEditedEquation } from './linear-format';

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<m:r><m:t>${text}</m:t></m:r>`;
const para = (...maths: string[]) =>
	`<m:oMathPara xmlns:m="${M}" xmlns:w="${W}"><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>${maths
		.map((math) => `<m:oMath>${math}</m:oMath>`)
		.join('')}</m:oMathPara>`;
// The first line keeps Word's exact spelling (attribute quotes, a comment), which must survive.
const FIRST = `<m:oMath><!-- a > b --><m:r><w:rPr><w:rFonts w:ascii='Cambria Math'/></w:rPr><m:t>a=b</m:t></m:r></m:oMath>`;

describe('display paragraphs with several equations', () => {
	it('counts the equations of a paragraph, and one for anything else', () => {
		expect(equationParts(para(run('a'), run('b'), run('c')))).toBe(3);
		expect(equationParts(para(run('a')))).toBe(1);
		expect(equationParts(`<m:oMath xmlns:m="${M}">${run('a')}</m:oMath>`)).toBe(1);
		expect(equationParts('<m:oMathPara')).toBe(1);
	});

	it('offers each equation for editing, though not the paragraph as a whole', () => {
		const source = para(run('a'), run('x+1'));
		expect(editableEquation(source, 'unicode')).toBeUndefined();
		expect(editableEquation(source, 'unicode', 0)?.text).toBe('a');
		expect(editableEquation(source, 'unicode', 1)?.text).toBe('x+1');
		expect(linearizeEquation(source, 'unicode', 1)).toBe('x+1');
		expect(editableEquation(source, 'unicode', 2)).toBeUndefined();
	});

	it('rewrites only the second equation, leaving the first and the paragraph byte-identical', () => {
		const source = `<m:oMathPara xmlns:m="${M}" xmlns:w="${W}"><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>${FIRST}<m:oMath>${run('x')}</m:oMath></m:oMathPara>`;
		const edit = editableEquation(source, 'unicode', 1)!;
		const math = writeEditedEquation({
			...edit,
			text: 'x/2',
			format: 'unicode',
			display: false,
		})!;
		const out = replaceEquationPart(source, 1, math)!;
		const before = source.slice(0, source.indexOf(FIRST) + FIRST.length);
		expect(out.startsWith(before)).toBe(true);
		expect(out.endsWith('</m:oMath></m:oMathPara>')).toBe(true);
		expect(out).toContain('<m:f>');
		// The namespaces the paragraph declares are not repeated on the rewritten equation.
		expect(out.slice(before.length)).toMatch(/^<m:oMath>/);
		expect(equationParts(out)).toBe(2);
		expect(editableEquation(out, 'unicode', 1)?.text).toBe('x/2');
	});

	it('refuses a part that does not exist or markup that is not an equation', () => {
		const source = para(run('a'), run('b'));
		expect(replaceEquationPart(source, 2, `<m:oMath xmlns:m="${M}"/>`)).toBeUndefined();
		expect(replaceEquationPart(source, 0, `<m:r xmlns:m="${M}"/>`)).toBeUndefined();
		expect(replaceEquationPart(`<m:oMath xmlns:m="${M}"/>`, 0, `<m:oMath/>`)).toBeUndefined();
	});

	it('keeps an equation read-only for tracked changes and mixed formatting, and says which', () => {
		const tracked = para(run('a'), `<w:ins w:id="1" w:author="A"><m:r><m:t>b</m:t></m:r></w:ins>`);
		const mixed = para(run('a'), `${run('b')}<m:r><w:rPr><w:b/></w:rPr><m:t>c</m:t></m:r>`);
		const box = para(run('a'), `<m:box><m:e>${run('x')}</m:e></m:box>`);
		for (const source of [tracked, mixed, box]) {
			expect(editableEquation(source, 'unicode', 0)).toBeDefined();
			expect(editableEquation(source, 'unicode', 1)).toBeUndefined();
			expect(equationReadOnlyReason(source, 'unicode', 0)).toBeUndefined();
		}
		expect(equationReadOnlyReason(tracked, 'unicode', 1)).toBe('tracked');
		expect(equationReadOnlyReason(mixed, 'unicode', 1)).toBe('formatting');
		expect(equationReadOnlyReason(box, 'unicode', 1)).toBe('content');
		const deleted = `<m:oMath xmlns:m="${M}" xmlns:w="${W}"><m:r><w:rPr><w:del w:id="2" w:author="A"/></w:rPr><m:t>x</m:t></m:r></m:oMath>`;
		expect(editableEquation(deleted, 'unicode')).toBeUndefined();
		expect(equationReadOnlyReason(deleted, 'unicode')).toBe('tracked');
		expect(equationReadOnlyReason('<m:oMath', 'unicode')).toBe('content');
	});
});

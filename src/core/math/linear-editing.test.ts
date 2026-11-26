import { describe, expect, it } from 'vitest';
import fixtures from './__fixtures__/word-equations.json';
import { canonicalMath } from './equation-canonical';
import { editableEquation, writeEditedEquation, type EquationInputFormat } from './linear-format';
import { parseOmml } from './omml-dom';
import type { OmmlNode } from './omml-node';
import { MAX_LINEAR_NESTING, unicodeMathToLatex } from './unicode-math';
import { buildUpEquation } from './linear-format';

const NS =
	' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const declare = (omml: string) => omml.replace(/^<(m:oMath(?:Para)?)/, `<$1${NS}`);
const font = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const run = (text: string) => `<m:r>${font}<m:t>${text}</m:t></m:r>`;
const para = (body: string) => declare(`<m:oMathPara>${body}</m:oMathPara>`);
const innerMath = (omml: string) =>
	(parseOmml(omml)['m:oMathPara'] as OmmlNode | undefined)?.['m:oMath'] as OmmlNode;

// Word 16 builds these from the linear text through its automation (OMath.BuildUp) and saves them;
// `src/core/scripts/math-word-equation-fixtures.ps1` regenerates the file.
const authored = fixtures.filter((f) => f.linear !== 'a ̂');

describe('editing equations Word wrote', () => {
	it.each(
		authored.flatMap((f) =>
			(['unicode', 'latex'] as const).map((format) => [f.linear, format, f.omml] as const),
		),
	)(
		'%s (%s) is offered as linear text and rewritten unchanged, run fonts kept',
		(_linear, format: EquationInputFormat, omml) => {
			const source = declare(omml);
			const edit = editableEquation(source, format)!;
			expect(edit).toBeDefined();
			const out = writeEditedEquation({ ...edit, format, display: true })!;
			expect(canonicalMath(innerMath(out))).toBe(canonicalMath(innerMath(source)));
			// Every run keeps Word's math font, and the markup declares the namespace it needs.
			const runs = out.match(/<m:r>/g)?.length ?? 0;
			expect(out.split(font).length - 1).toBe(runs);
			expect(out).toMatch(/^<m:oMathPara xmlns:m="[^"]+" xmlns:w="[^"]+">/);
		},
	);

	it('keeps the paragraph justification of a display equation', () => {
		const source = para(
			`<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>${run('x')}</m:oMath>`,
		);
		const edit = editableEquation(source, 'unicode')!;
		const out = writeEditedEquation({ ...edit, text: 'x+1', format: 'unicode', display: true })!;
		expect(out).toContain('<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>');
	});

	it.each([
		[
			'two equations in one display paragraph',
			para(`<m:oMath>${run('a')}</m:oMath><m:oMath>${run('b')}</m:oMath>`),
		],
		[
			'runs formatted differently',
			para(`<m:oMath>${run('a')}<m:r><w:rPr><w:b/></w:rPr><m:t>b</m:t></m:r></m:oMath>`),
		],
		[
			'a bold math style',
			para(`<m:oMath><m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>x</m:t></m:r></m:oMath>`),
		],
		[
			'a manual line break',
			para(`<m:oMath><m:r><m:rPr><m:brk/></m:rPr><m:t>=</m:t></m:r></m:oMath>`),
		],
		['a box wrapper', para(`<m:oMath><m:box><m:e>${run('x')}</m:e></m:box></m:oMath>`)],
		[
			'a phantom (LaTeX keeps it)',
			para(`<m:oMath><m:phant><m:e>${run('x')}</m:e></m:phant></m:oMath>`),
		],
		[
			'a group character other than a brace',
			para(
				`<m:oMath><m:groupChr><m:groupChrPr><m:chr m:val="→"/></m:groupChrPr><m:e>${run('x')}</m:e></m:groupChr></m:oMath>`,
			),
		],
		[
			'an accent that is not a combining mark',
			para(
				`<m:oMath><m:acc><m:accPr><m:chr m:val="~"/></m:accPr><m:e>${run('x')}</m:e></m:acc></m:oMath>`,
			),
		],
		[
			'a custom bracket separator (LaTeX keeps it)',
			para(
				`<m:oMath><m:d><m:dPr><m:sepChr m:val=","/></m:dPr><m:e>${run('a')}</m:e><m:e>${run('b')}</m:e></m:d></m:oMath>`,
			),
		],
		[
			'limits above an integral (LaTeX keeps it)',
			para(
				`<m:oMath><m:nary><m:naryPr><m:limLoc m:val="undOvr"/></m:naryPr><m:sub>${run('0')}</m:sub><m:sup>${run('1')}</m:sup><m:e>${run('x')}</m:e></m:nary></m:oMath>`,
			),
		],
		[
			'left-aligned matrix columns',
			para(
				`<m:oMath><m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="left"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e>${run('a')}</m:e></m:mr></m:m></m:oMath>`,
			),
		],
		[
			'an alignment point',
			para(`<m:oMath><m:r><m:rPr><m:aln/></m:rPr><m:t>=</m:t></m:r></m:oMath>`),
		],
		[
			'a script style',
			para(
				`<m:oMath><m:r><m:rPr><m:scr m:val="double-struck"/></m:rPr><m:t>R</m:t></m:r></m:oMath>`,
			),
		],
		['unreadable markup', '<m:oMath'],
	])('keeps an equation with %s read-only', (_name, source) => {
		expect(editableEquation(source, 'unicode')).toBeUndefined();
		if (!_name.includes('LaTeX keeps it'))
			expect(editableEquation(source, 'latex')).toBeUndefined();
	});
});

describe('linear format nesting', () => {
	it.each(['(', '〖', '√', '∑', '■(', '¯'])(
		'stops at %s nested past the limit with an error',
		(open) => {
			const text = open.repeat(16_000);
			expect(() => unicodeMathToLatex(text)).toThrow(/nesting exceeds/);
			expect(buildUpEquation(text, 'unicode', false)).toBeUndefined();
		},
	);

	it('builds nesting up to the limit', () => {
		const depth = Math.floor(MAX_LINEAR_NESTING / 4);
		expect(
			buildUpEquation(`${'('.repeat(depth)}x${')'.repeat(depth)}`, 'unicode', false),
		).toContain('<m:d>');
	});
});

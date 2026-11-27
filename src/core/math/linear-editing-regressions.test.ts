// Regressions found by probing linear editing for losses: each case either rewrites the equation
// exactly or is refused (read-only, or Convert declines), never changed silently.
import { describe, expect, it } from 'vitest';
import { canonicalMath, coalesceRuns } from './equation-canonical';
import {
	buildUpEquation,
	editableEquation,
	linearDisplayText,
	writeEditedEquation,
	type EquationInputFormat,
} from './linear-format';
import { convertLatexToOmml } from './latex-to-omml';
import { parseOmml } from './omml-dom';
import type { OmmlNode } from './omml-node';

const NS =
	' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const F = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const RED =
	'<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:color w:val="FF0000"/></w:rPr>';
const R = (t: string, pre = '') => `<m:r>${pre}${F}<m:t>${t}</m:t></m:r>`;
const M = (body: string) => `<m:oMathPara${NS}><m:oMath>${body}</m:oMath></m:oMathPara>`;
const inner = (omml: string): OmmlNode => {
	const tree = parseOmml(omml);
	return ((tree['m:oMathPara'] as OmmlNode | undefined)?.['m:oMath'] ??
		tree['m:oMath']) as OmmlNode;
};
const sameAs = (a: string, b: string) => canonicalMath(inner(a)) === canonicalMath(inner(b));

/** Converts to Linear and back to Professional; true when it is refused or exact. */
function linearRoundTrip(omml: string, format: EquationInputFormat): boolean {
	const edit = editableEquation(omml, format);
	if (!edit) return true;
	const linear = writeEditedEquation({ ...edit, format, display: true, linear: true });
	if (!edit.canLinear) return linear === undefined;
	if (!linear) return false;
	const text = linearDisplayText(linear);
	const back = text === undefined ? linear : buildUpEquation(text, 'unicode', true);
	return Boolean(back && sameAs(back, omml));
}

const PRESERVE = {
	phantomHidden: M(
		`<m:phant><m:phantPr><m:show m:val="0"/><m:ctrlPr>${F}</m:ctrlPr></m:phantPr><m:e>${R('x')}</m:e></m:phant>`,
	),
	commaSeparator: M(
		`<m:d><m:dPr><m:sepChr m:val=","/><m:ctrlPr>${F}</m:ctrlPr></m:dPr><m:e>${R('a')}</m:e><m:e>${R('b')}</m:e></m:d>`,
	),
	limitsAboveIntegral: M(
		`<m:nary><m:naryPr><m:limLoc m:val="undOvr"/></m:naryPr><m:sub>${R('0')}</m:sub><m:sup>${R('1')}</m:sup><m:e>${R('x')}</m:e></m:nary>`,
	),
	textBeforeFraction: M(`${R('a')}<m:f><m:num>${R('b')}</m:num><m:den>${R('c')}</m:den></m:f>`),
	textBeforeScript: M(`<m:sSup><m:e>${R('x')}</m:e><m:sup>${R('2')}</m:sup></m:sSup>${R('y')}`),
	textBeforeFunction: M(
		`${R('a')}<m:func><m:fName>${R('sin', '<m:rPr><m:sty m:val="p"/></m:rPr>')}</m:fName><m:e>${R('x')}</m:e></m:func>`,
	),
	emptyShownDegree: M(
		`<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg/><m:e>${R('x')}</m:e></m:rad>`,
	),
};

describe('Convert to Linear', () => {
	it.each(
		Object.entries(PRESERVE).flatMap(([name, omml]) => [
			[name, 'unicode', omml] as const,
			[name, 'latex', omml] as const,
		]),
	)('%s in %s input is refused or exact', (_name, format, omml) => {
		expect(linearRoundTrip(omml, format)).toBe(true);
	});

	it('writes a space between text and a fraction after it, as Word writes `a b/c`', () => {
		const edit = editableEquation(PRESERVE.textBeforeFraction, 'unicode')!;
		expect(edit.text).toBe('a b/c');
	});

	it('refuses Linear for a hidden phantom typed in LaTeX', () => {
		const edit = editableEquation(PRESERVE.phantomHidden, 'latex')!;
		expect(edit.canLinear).toBe(false);
		expect(
			writeEditedEquation({ ...edit, format: 'latex', display: true, linear: true }),
		).toBeUndefined();
	});
});

describe('runs and control characters', () => {
	it.each([
		['a symbol', `<m:r>${F}<w:sym w:font="Symbol" w:char="F061"/><m:t>b</m:t></m:r>`],
		['a break', `<m:r>${F}<w:br/><m:t>b</m:t></m:r>`],
		['a drawing', `<m:r>${F}<w:drawing/><m:t>b</m:t></m:r>`],
		['deleted text', `<m:r>${F}<w:delText>q</w:delText><m:t>b</m:t></m:r>`],
		['an attribute', `<m:r w:rsidR="00AB">${F}<m:t>b</m:t></m:r>`],
		['a tracked insertion', `<w:ins w:id="1">${R('b')}</w:ins>`],
	])('keeps a run with %s read-only and never merges it', (_name, run) => {
		const omml = M(`<m:f><m:num>${R('a')}${run}</m:num><m:den>${R('c')}</m:den></m:f>`);
		for (const format of ['unicode', 'latex'] as const)
			expect(editableEquation(omml, format)).toBeUndefined();
		const merged = coalesceRuns({
			'm:r': [{ 'm:t': 'a' }, { 'w:sym': { '@_char': 'F061' }, 'm:t': 'b' }],
		});
		expect(JSON.stringify(merged)).toContain('F061');
		expect(JSON.stringify(merged)).toContain('"a"');
	});

	it('keeps control characters with more than the math font read-only', () => {
		const red = M(
			`<m:f><m:fPr><m:ctrlPr>${RED}</m:ctrlPr></m:fPr><m:num>${R('a')}</m:num><m:den>${R('b')}</m:den></m:f>`,
		);
		const tracked = M(
			`<m:f><m:fPr><m:ctrlPr><w:ins w:id="1" w:author="x">${F}</w:ins></m:ctrlPr></m:fPr><m:num>${R('a')}</m:num><m:den>${R('b')}</m:den></m:f>`,
		);
		expect(editableEquation(red, 'unicode')).toBeUndefined();
		expect(editableEquation(tracked, 'unicode')).toBeUndefined();
	});

	it('writes the math font back on the control characters', () => {
		const ctrl = `<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>`;
		const omml = M(
			`<m:f><m:fPr>${ctrl}</m:fPr><m:num>${R('a')}</m:num><m:den>${R('b')}</m:den></m:f>`,
		);
		const edit = editableEquation(omml, 'unicode')!;
		const out = writeEditedEquation({ ...edit, text: 'a/c', format: 'unicode', display: true })!;
		expect(out).toContain(`<m:fPr>${ctrl}</m:fPr>`);
	});

	it('refuses other namespaces in run formatting instead of writing unreadable markup', () => {
		const ns = `${NS} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
		const shadow =
			'<w:rPr><w:rFonts w:ascii="Cambria Math"/><w14:shadow w14:blurRad="38100"/></w:rPr>';
		const omml = `<m:oMathPara${ns}><m:oMath><m:r>${shadow}<m:t>x</m:t></m:r></m:oMath></m:oMathPara>`;
		expect(editableEquation(omml, 'unicode')).toBeUndefined();
	});
});

describe('on and off properties', () => {
	it('tells a hidden phantom from a shown one, and LaTeX phantoms are hidden', () => {
		const hidden = M(
			`<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr><m:e>${R('x')}</m:e></m:phant>`,
		);
		const shown = M(`<m:phant><m:e>${R('x')}</m:e></m:phant>`);
		expect(sameAs(hidden, shown)).toBe(false);
		const phantom = convertLatexToOmml('\\phantom{x}')['m:oMathPara'] as OmmlNode;
		expect(JSON.stringify(phantom)).toContain('"m:show":{"@_val":"0"}');
	});

	it('treats an explicit off like an absent default-off flag', () => {
		const off = M(
			`<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg/><m:e>${R('x')}</m:e></m:rad>`,
		);
		const absent = M(`<m:rad><m:deg/><m:e>${R('x')}</m:e></m:rad>`);
		expect(sameAs(off, absent)).toBe(true);
	});
});

describe('text at the top of an equation', () => {
	it.each(['E=m', 'a+b', '2x', 'α+β'])(
		'Word-written %s is editable and saved as one run',
		(text) => {
			const omml = M(R(text));
			for (const format of ['unicode', 'latex'] as const) {
				const edit = editableEquation(omml, format)!;
				expect(edit).toBeDefined();
				const out = writeEditedEquation({ ...edit, format, display: true })!;
				expect(out.match(/<m:r>/g)).toHaveLength(1);
			}
		},
	);

	it("Word's E=mc² stays editable", () => {
		const omml = M(`${R('E=m')}<m:sSup><m:e>${R('c')}</m:e><m:sup>${R('2')}</m:sup></m:sSup>`);
		expect(editableEquation(omml, 'unicode')?.text).toBe('E=m c^2');
	});

	it('saves km\\/h as one run that later reads as Professional text', () => {
		const out = buildUpEquation('km\\/h', 'unicode', false)!;
		expect(out.match(/<m:r>/g)).toHaveLength(1);
		expect(linearDisplayText(out)).toBeUndefined();
	});
});

describe('Linear of what its text can hold', () => {
	it.each([
		['an accent', M(`<m:acc><m:accPr><m:chr m:val="̃"/></m:accPr><m:e>${R('a')}</m:e></m:acc>`)],
		['text with a space', M(`<m:r>${F}<m:t xml:space="preserve">a b</m:t></m:r>`)],
		['a function after text', PRESERVE.textBeforeFunction],
	])('converts %s to Linear and back exactly', (_name, omml) => {
		for (const format of ['unicode', 'latex'] as const) {
			const edit = editableEquation(omml, format)!;
			expect(edit.canLinear).toBe(true);
			expect(writeEditedEquation({ ...edit, format, display: true, linear: true })).toBeDefined();
			expect(linearRoundTrip(omml, format)).toBe(true);
		}
	});
});

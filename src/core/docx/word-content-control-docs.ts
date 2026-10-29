// The document of `scripts/docx-word-content-control-probe.ps1`, rebuilt in TypeScript so the layout
// can be compared with what Word measured (`__fixtures__/word-content-control/line-probe.json`).
// Keep the two in step: every case and run here is the script's.
import { packProbe } from './word-spacing-docs';

const run = (text: string, rPr = '') =>
	`<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const gothic =
	'<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/>';
const gothicLatin = '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/>';
const segoe = '<w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol"/>';
const BOX = '☒';
const EMPTY = '☐';
const check = (state: string, glyph: string, rPr: string) =>
	`<w:sdt><w:sdtPr><w:id w:val="1001"/><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="${state}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent>${run(glyph, rPr)}</w:sdtContent></w:sdt>`;

/** Case name to the runs of its paragraph, in the script's order. */
export const contentControlCases: Record<string, string> = {
	plain: run('Plain: ok'),
	gothicsym: run('Gothic: ') + run(BOX, gothic),
	gothictxt: run('Gothic: ') + run('abc', gothic),
	gothiclat: run('Gothic: ') + run('abc', gothicLatin),
	gothiconly: run(BOX, gothic),
	cbchecked: run('Agree: ') + check('1', BOX, gothic),
	cbunchecked: run('Agree: ') + check('0', EMPTY, gothic),
	cbcourier: run('Agree: ') + check('1', BOX, ''),
	text:
		run('Name: ') +
		`<w:sdt><w:sdtPr><w:id w:val="1002"/><w:text/></w:sdtPr><w:sdtContent>${run('Ada')}</w:sdtContent></w:sdt>`,
	dropdown:
		run('Colour: ') +
		`<w:sdt><w:sdtPr><w:id w:val="1003"/><w:dropDownList><w:listItem w:displayText="Red" w:value="Red"/></w:dropDownList></w:sdtPr><w:sdtContent>${run('Red')}</w:sdtContent></w:sdt>`,
	symcourier: run('Sym: ') + run(BOX),
	symsegoe: run('Sym: ') + run(BOX, segoe),
};

/** The probe document: a label paragraph (`L<case>`) before each case paragraph, then `Lend`. */
export function contentControlDocx(): Promise<Uint8Array> {
	const body =
		Object.entries(contentControlCases)
			.map(([name, runs]) => `<w:p><w:r><w:t>L${name}</w:t></w:r></w:p><w:p>${runs}</w:p>`)
			.join('') + '<w:p><w:r><w:t>Lend</w:t></w:r></w:p>';
	return packProbe(body, undefined, 15, false);
}

import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { latexToOmmlXml } from '../../math/index';
import {
	createEditSession,
	createWorkbook,
	loadXlsx,
	pictureAnchorAt,
	saveXlsx,
	type Workbook,
} from '../index';

const markup = (latex: string) => ({ omml: latexToOmmlXml(latex, true)!, fallback: latex });
const fixture = () =>
	readFile(new URL('../__fixtures__/excel-equation-native.xlsx', import.meta.url));
const drawingXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('xl/drawings/drawing1.xml')!.async('string');
const shape = (workbook: Workbook, index = 0) => {
	const drawing = workbook.sheets[0]?.drawings[index];
	return drawing?.kind === 'unsupported' ? drawing.shape : undefined;
};

it('writes a new equation text box with the native choice and linear fallback', async () => {
	const book = createWorkbook();
	const session = createEditSession(book);
	const index = session.addShape(0, pictureAnchorAt({ row: 3, col: 2 }, 150, 55), {
		kind: 'equation',
		name: 'Equation',
		equation: markup('\\frac{a}{b}'),
	});
	expect(session.setShapeEquation(0, index, 0, markup('x^2'))).toBe(true);
	expect(session.undo()).toBe(true);
	const bytes = await saveXlsx(book);
	const xml = await drawingXml(bytes);
	expect(xml).toContain('<mc:AlternateContent');
	expect(xml).toContain('Requires="a14"><xdr:sp>');
	expect(xml).toContain('<a14:m><m:oMathPara');
	expect(xml).toContain('<m:f>');
	expect(xml).toMatch(/<mc:Fallback>.*<a:t>\\frac\{a\}\{b\}<\/a:t>.*<\/mc:Fallback>/s);
	expect(shape(await loadXlsx(bytes))?.equationParagraphs).toMatchObject([
		[{ kind: 'equation', fallback: '\\frac{a}{b}' }],
	]);
});

it('keeps the native Excel shape and fallback until the equation is edited', async () => {
	const native = await fixture();
	const book = await loadXlsx(native);
	expect(shape(book)).toMatchObject({
		textBox: true,
		equationParagraphs: [[{ kind: 'equation', fallback: '"Type equation here."' }]],
	});
	expect(await drawingXml(await saveXlsx(book))).toEqual(await drawingXml(native));
	const session = createEditSession(book);
	session.setTextBoxText(0, 0, 'Do not overwrite Office Math through the text command');
	expect(session.setShapeEquation(0, 0, 0, markup('\\sqrt{2}'))).toBe(true);
	const editedBytes = await saveXlsx(book);
	const edited = await drawingXml(editedBytes);
	expect(edited).toContain('<m:rad>');
	expect(edited).toContain('\\sqrt{2}');
	expect(edited).toContain('a16:creationId');
	expect(edited).not.toContain('Type equation here.');
	expect(shape(await loadXlsx(editedBytes))?.equationParagraphs?.[0]?.[0]).toMatchObject({
		fallback: '\\sqrt{2}',
	});
	expect(session.undo()).toBe(true);
	expect(await drawingXml(await saveXlsx(book))).toEqual(await drawingXml(native));
});

it('edits a second equation while preserving the first equation and the adjacent text', async () => {
	const native = await fixture();
	const zip = await JSZip.loadAsync(native);
	const original = await drawingXml(native);
	const run = (text: string) => `<a:r><a:t>${text}</a:t></a:r>`;
	const note = run(' and ');
	zip.file(
		'xl/drawings/drawing1.xml',
		original
			.replace('</a14:m>', `</a14:m>${note}<a14:m>${markup('x^2').omml}</a14:m>`)
			.replace(
				'"Type equation here."</a:t></a:r>',
				`"Type equation here."</a:t></a:r>${note}${run('x^2')}`,
			),
	);
	const book = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	const session = createEditSession(book);
	expect(session.setShapeEquation(0, 0, 1, markup('\\frac{a}{b}'))).toBe(true);
	const drawing = shape(await loadXlsx(await saveXlsx(book)));
	expect(drawing?.equationParagraphs).toMatchObject([
		[
			{ kind: 'equation', fallback: '"Type equation here."' },
			{ kind: 'text', text: ' and ' },
			{ kind: 'equation', fallback: '\\frac{a}{b}', omml: expect.stringContaining('<m:f>') },
		],
	]);
});

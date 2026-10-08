import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, shape, rectangle, section } from './test-fixtures';

const command = { type: 'format-text' as const, pageId: '0', shapeId: '1' };
const row = (index: number, cells: string) => `<Row IX="${index}">${cells}</Row>`;
const fonts =
	'<FaceNames><FaceName ID="0" Name="Arial"/><FaceName ID="1" Name="Calibri"/></FaceNames>';
const source = (
	rows: string,
	text = '<cp IX="0"/>First<cp IX="1"/>Second',
	document = fonts,
	attrs = '',
) =>
	fixture({
		document,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape(
					'1',
					cell('Width', 4) + cell('Height', 2) + rectangle + rows + `<Text>${text}</Text>`,
					attrs,
				)}</Shapes>`,
			},
		],
		edit: (zip) => zip.file('unknown/rich.bin', new Uint8Array([0, 3, 255])),
	});
const xml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');
const read = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;

describe('rich source-backed formatting', () => {
	it('formats every character row while preserving rowwise unrelated styles and native markers', async () => {
		const text = '<pp IX="0"/><tp IX="0"/><cp IX="0"/>First<cp IX="1"/>Second\n<pp IX="1"/>Third';
		const rows =
			section(
				'Character',
				row(0, cell('Style', 2) + cell('Size', 0.2)) +
					row(1, cell('Style', 12) + cell('Size', 0.3)),
			) +
			section(
				'Paragraph',
				row(0, cell('Bullet', 2)) + row(1, cell('Bullet', 7) + cell('BulletStr', '★')),
			);
		const bytes = await source(rows, text);
		const before = bytes.slice();
		const saved = await editVsdx(bytes, [
			{
				...command,
				bold: true,
				fontColor: '#123456',
				fontSize: 18,
				fontFamily: 'Calibri',
				strikethrough: true,
				bullets: true,
				indentLeft: 36,
				horizontalAlign: 'justify',
			},
		]);
		expect(bytes).toEqual(before);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(await xml(saved.bytes)).toContain(`<Text>${text}</Text>`);
		expect(await xml(saved.bytes)).toContain('N="Style" V="3"');
		expect(await xml(saved.bytes)).toContain('N="Style" V="13"');
		const parsed = await read(saved.bytes);
		expect(parsed.text.runs.map((run) => [run.italic, run.underline])).toEqual([
			[true, false],
			[false, true],
			[false, true],
		]);
		for (const run of parsed.text.runs)
			expect(run).toMatchObject({
				bold: true,
				fontFamily: 'Calibri',
				fontSize: 0.25,
				color: '#123456',
				strikethrough: true,
			});
		expect(parsed.text.paragraphs?.map((p) => p.bullet?.text)).toEqual(['◆', '★']);
		for (const paragraph of parsed.text.paragraphs ?? [])
			expect(paragraph).toMatchObject({ indentLeft: 0.5, horizontalAlign: 'justify' });
		const reopened = await editVsdx(saved.bytes, [
			{
				...command,
				bold: true,
				fontColor: '#123456',
				fontSize: 18,
				fontFamily: 'Calibri',
				strikethrough: true,
				bullets: true,
				indentLeft: 36,
				horizontalAlign: 'justify',
			},
		]);
		expect(reopened.changedParts).toEqual([]);
		expect(reopened.bytes).toEqual(saved.bytes);
		const beforeZip = await JSZip.loadAsync(bytes),
			afterZip = await JSZip.loadAsync(saved.bytes);
		for (const path of Object.keys(beforeZip.files))
			if (!beforeZip.files[path]!.dir && path !== 'visio/pages/page1.xml')
				expect(await afterZip.file(path)!.async('uint8array')).toEqual(
					await beforeZip.file(path)!.async('uint8array'),
				);
	});
	it('uses corresponding inherited rows before default rows and never inherits a local sibling', async () => {
		const document =
			fonts +
			'<StyleSheets><StyleSheet ID="0">' +
			section('Character', row(0, cell('Style', 2)) + row(1, cell('Style', 4))) +
			'</StyleSheet><StyleSheet ID="2" TextStyle="0">' +
			section('Character', row(0, cell('Style', 8))) +
			'</StyleSheet></StyleSheets>';
		const bytes = await source(
			section(
				'Character',
				row(0, cell('Style', 16)) + row(1, cell('Size', 0.2)) + row(2, cell('Size', 0.3)),
			),
			'<cp IX="0"/>A<cp IX="1"/>B<cp IX="2"/>C',
			document,
			'TextStyle="2"',
		);
		const saved = await editVsdx(bytes, [{ ...command, bold: true }]);
		const xmlText = await xml(saved.bytes);
		for (const bits of [17, 5, 9]) expect(xmlText).toContain(`N="Style" V="${bits}"`);
		expect(
			(await read(saved.bytes)).text.runs.map((run) => [run.bold, run.italic, run.underline]),
		).toEqual([
			[true, false, false],
			[true, false, true],
			[true, false, false],
		]);
	});
	it('materializes inherited rows without rewriting document styles or changing unrelated row properties', async () => {
		const document =
			fonts +
			'<StyleSheets><StyleSheet ID="2">' +
			section(
				'Character',
				row(0, cell('Style', 2) + cell('Size', 0.2)) + row(1, cell('Style', 4) + cell('Size', 0.4)),
			) +
			'</StyleSheet></StyleSheets>';
		const bytes = await source('', undefined, document, 'TextStyle="2"');
		const saved = await editVsdx(bytes, [{ ...command, bold: true }]);
		expect(
			(await read(saved.bytes)).text.runs.map((run) => [
				run.fontSize,
				run.bold,
				run.italic,
				run.underline,
			]),
		).toEqual([
			[0.2, true, true, false],
			[0.4, true, false, true],
		]);
		const zipBefore = await JSZip.loadAsync(bytes),
			zipAfter = await JSZip.loadAsync(saved.bytes);
		expect(await zipAfter.file('visio/document.xml')!.async('string')).toEqual(
			await zipBefore.file('visio/document.xml')!.async('string'),
		);
	});
	it('preserves cached fields, tab rows and all cp/pp/tp/fld markup while changing surrounding runs', async () => {
		const text = '<pp IX="0"/><cp IX="0"/>Value:<fld IX="0">2</fld><tp IX="1"/>\t<cp IX="1"/>end';
		const rows =
			section('Character', row(0, cell('Style', 0)) + row(1, cell('Style', 2))) +
			section('Tabs', row(1, cell('Position', 1))) +
			section('Field', row(0, cell('Value', 2, '1+1') + cell('Format', 0)));
		const bytes = await source(rows, text);
		const saved = await editVsdx(bytes, [{ ...command, underline: true }]);
		expect(await xml(saved.bytes)).toContain(`<Text>${text}</Text>`);
		expect(await xml(saved.bytes)).toContain(
			section('Field', row(0, cell('Value', 2, '1+1') + cell('Format', 0))),
		);
		expect((await read(saved.bytes)).text.plainText).toBe('Value:2\tend');
		expect((await read(saved.bytes)).text.runs.every((run) => run.underline)).toBe(true);
	});
});

import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, shape, rectangle, section } from './test-fixtures';
import { visioFormattingShape, visioTextFormattingState } from './ui/formatting';

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

describe('rich formatting admission', () => {
	it.each(['GUARD(0)', 'SETATREF(Width)', 'Width/10', 'Inh'])(
		'refuses a protected later row atomically: %s',
		async (formula) => {
			const bytes = await source(
				section('Character', row(0, cell('Size', 0.2)) + row(1, cell('Size', 0.3, formula))),
			);
			const before = bytes.slice();
			await expect(editVsdx(bytes, [{ ...command, fontSize: 18 }])).rejects.toThrow(
				/formula|protected|overwritten|transform/i,
			);
			expect(bytes).toEqual(before);
		},
	);
	it('refuses a guarded inherited nonzero row despite an unprotected style default', async () => {
		const document =
			fonts +
			'<StyleSheets><StyleSheet ID="2">' +
			section('Character', row(0, cell('Style', 0)) + row(1, cell('Style', 2, 'GUARD(2)'))) +
			'</StyleSheet></StyleSheets>';
		await expect(
			editVsdx(await source('', undefined, document, 'TextStyle="2"'), [
				{ ...command, bold: true },
			]),
		).rejects.toThrow(/guard|transform/i);
	});
	it.each(['Character.1.Size', 'Char.Size[2]'])(
		'refuses affected mixed-row field dependencies: %s',
		async (name) => {
			const bytes = await source(
				section('Character', row(0, '') + row(1, cell('Size', 0.3))) +
					section('Field', row(0, cell('Value', 2, `${name}*2`))),
				'<cp IX="1"/>A<fld IX="0">2</fld>',
			);
			await expect(editVsdx(bytes, [{ ...command, fontSize: 18 }])).rejects.toThrow(/dependenc/i);
		},
	);
	it.each([
		['duplicate', section('Character', '<Row IX="0"/><Row IX="0"/>'), '<cp IX="0"/>A'],
		['noncanonical', section('Character', '<Row IX="01"/>'), '<cp IX="0"/>A'],
		['deleted', section('Character', '<Row IX="1" Del="1"/>'), '<cp IX="0"/>A'],
		['named', section('Character', '<Row IX="0" N="named"/>'), 'A'],
		['missing', '', '<cp/>A'],
		['dangling', '', '<cp IX="1"/>A'],
		['nested', '', '<pp IX="0"><cp IX="0"/></pp>A'],
		['field row', '', '<fld IX="0">2</fld>'],
		['field markup', section('Field', row(0, cell('Value', 2))), '<fld IX="0"><cp IX="0"/>2</fld>'],
		['namespace', '', '<cp xmlns="urn:other" IX="0"/>A'],
	])('refuses malformed %s text without changing source bytes', async (_reason, rows, text) => {
		const bytes = await source(rows, text),
			original = bytes.slice();
		await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow(/format|row|text/i);
		expect(bytes).toEqual(original);
	});
	it('admits mixed scene runs with mixed state and toggles them all on', async () => {
		const document = await parseVsdx(
			await source(section('Character', row(0, cell('Style', 1)) + row(1, cell('Style', 2)))),
		);
		const shape = document.pages[0]!.shapes[0]!;
		expect(visioFormattingShape(document.pages[0]!, '1')).toBe(shape);
		expect(visioTextFormattingState([shape])).toMatchObject({ bold: false, italic: false });
	});
	it('keeps enabled child rows when a disabled ancestor suppresses inherited rows', async () => {
		const document =
			fonts +
			'<StyleSheets><StyleSheet ID="0">' +
			cell('EnableTextProps', 0) +
			section('Character', row(0, cell('Style', 2)) + row(1, cell('Style', 4))) +
			'</StyleSheet><StyleSheet ID="2" TextStyle="0">' +
			section('Character', row(0, cell('Style', 8))) +
			'</StyleSheet></StyleSheets>';
		const bytes = await source(
			section('Character', row(1, cell('Size', 0.2))),
			undefined,
			document,
			'TextStyle="2"',
		);
		const saved = await editVsdx(bytes, [{ ...command, bold: true }]);
		expect(
			(await read(saved.bytes)).text.runs.map((run) => [run.bold, run.italic, run.underline]),
		).toEqual([
			[true, false, false],
			[true, false, false],
		]);
		expect(await xml(saved.bytes)).toContain('N="Style" V="9"');
	});
	it.each([
		'THEMEGUARD(THEMEVAL(&quot;TextColor&quot;,SETATREFEXPR(Width)))',
		'THEMEGUARD(THEMEVAL(&quot;TextColor&quot;,GUARD(0)))',
		'THEMEVAL(&quot;TextColor&quot;,UNKNOWN())',
		'THEMEGUARD(THEMEVAL(&quot;TextColor&quot;,Sheet.1!Width))',
		'THEMEGUARD(THEMEVAL(&quot;TextColor&quot;,0,1))',
		'THEMEGUARD(0,1)',
		'THEMEVAL(99)',
	])('does not admit unsafe or malformed nested theme formatting: %s', async (formula) => {
		const bytes = await source(
			section(
				'Character',
				row(0, cell('Color', 'Themed', formula)) + row(1, cell('Color', '#abcdef')),
			),
		);
		await expect(editVsdx(bytes, [{ ...command, fontColor: '#123456' }])).rejects.toThrow(
			/protected|transform|formula/i,
		);
	});
	it('changes only vertical alignment without needing local text or valid text rows', async () => {
		const document =
			fonts +
			'<StyleSheets><StyleSheet ID="2">' +
			section('Character', '<Row IX="0" Del="1"/>') +
			'</StyleSheet></StyleSheets>';
		const bytes = await fixture({
			document,
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape(
						'1',
						cell('Width', 4) +
							cell('Height', 2) +
							rectangle +
							section('Paragraph', '<Row IX="0" N="custom"/>'),
						'TextStyle="2"',
					)}</Shapes>`,
				},
			],
		});
		const saved = await editVsdx(bytes, [{ ...command, verticalAlign: 'top' }]);
		expect(await xml(saved.bytes)).toContain('N="VerticalAlign" V="0"');
		expect(await xml(saved.bytes)).toContain('<Row IX="0" N="custom"/>');
		await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow(/format/i);
	});
	it('does not require unrelated Character row admission for paragraph formatting', async () => {
		const bytes = await source(
			section('Character', '<Row IX="0" N="custom"/>') +
				section('Paragraph', row(0, cell('HorzAlign', 0))),
			'<cp IX="0"/>A',
		);
		const saved = await editVsdx(bytes, [{ ...command, horizontalAlign: 'right' }]);
		expect((await read(saved.bytes)).text.paragraphs?.[0]?.horizontalAlign).toBe('right');
		expect(await xml(saved.bytes)).toContain('<Row IX="0" N="custom"/>');
	});
	it('does not require unrelated Paragraph row admission for character formatting', async () => {
		const bytes = await source(
			section('Character', row(0, cell('Style', 0))) +
				section('Paragraph', '<Row IX="0" N="custom"/>'),
			'<pp IX="0"/>A',
		);
		const saved = await editVsdx(bytes, [{ ...command, bold: true }]);
		expect((await read(saved.bytes)).text.runs[0]?.bold).toBe(true);
		expect(await xml(saved.bytes)).toContain('<Row IX="0" N="custom"/>');
	});
	it('keeps unused protected character rows protected under the stored-row whole-shape contract', async () => {
		const bytes = await source(
			section('Character', row(0, cell('Style', 0)) + row(1, cell('Style', 2, 'GUARD(2)'))),
			'<cp IX="0"/>Only default row',
		);
		await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow(/transform|guard/i);
	});
	it('retains inherited units when proving partially overridden scalar styles', async () => {
		const document =
			fonts +
			'<StyleSheets><StyleSheet ID="2">' +
			section('Character', row(0, '<Cell N="Style" V="2" U="IN"/>')) +
			'</StyleSheet></StyleSheets>';
		const bytes = await source(
			section('Character', row(0, cell('Style', 0)) + row(1, cell('Style', 4))),
			undefined,
			document,
			'TextStyle="2"',
		);
		await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow(/scalar/i);
	});
	it('bounds row expansion before changing any package bytes', async () => {
		const bytes = await source(
			section('Character', Array.from({ length: 10_001 }, (_, index) => row(index, '')).join('')),
			'<cp IX="0"/>A',
		);
		const original = bytes.slice();
		await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow(/row limit/i);
		expect(bytes).toEqual(original);
	});
	it('proves native delegated style caches before formatting inherited rich rows', async () => {
		const document =
			fonts +
			'<StyleSheets><StyleSheet ID="0">' +
			section('Character', row(0, cell('Style', 2) + cell('ColorTrans', 0))) +
			'</StyleSheet>' +
			'<StyleSheet ID="6" TextStyle="0">' +
			section(
				'Character',
				row(
					0,
					cell('Style', 2, 'Inh') +
						cell('Color', 'Themed', 'THEMEVAL()') +
						cell('ColorTrans', 0, 'Inh'),
				),
			) +
			'</StyleSheet><StyleSheet ID="3" TextStyle="6"/></StyleSheets>';
		const bytes = await source(
			section('Character', row(0, cell('Style', 0)) + row(1, cell('Size', 0.2))),
			undefined,
			document,
			'TextStyle="3"',
		);
		const saved = await editVsdx(bytes, [{ ...command, bold: true, fontColor: '#123456' }]);
		expect(
			(await read(saved.bytes)).text.runs.map((run) => [run.bold, run.italic, run.color]),
		).toEqual([
			[true, false, '#123456'],
			[true, true, '#123456'],
		]);
	});
	it.each([
		['stale', 4, 'TextStyle="0"'],
		['missing ancestor', 2, ''],
	])(
		'refuses %s inherited delegation rather than trusting rich-run caches',
		async (_reason, cache, attr) => {
			const document =
				fonts +
				'<StyleSheets><StyleSheet ID="0">' +
				section('Character', row(0, cell('Style', 2))) +
				'</StyleSheet>' +
				`<StyleSheet ID="6" ${attr}>` +
				section('Character', row(0, cell('Style', cache, 'Inh'))) +
				'</StyleSheet></StyleSheets>';
			const bytes = await source(
				section('Character', row(0, cell('Style', 0)) + row(1, cell('Size', 0.2))),
				undefined,
				document,
				'TextStyle="6"',
			);
			await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow(
				/inherited formatting/i,
			);
		},
	);
});

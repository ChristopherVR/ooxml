import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx, type VisioFormatEdit } from './edit';
import { parseVsdx } from './parser';
import { snapshotEdits } from './ui/edit-commands';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import {
	visioFormattingShape,
	visioStyleFormattingShape,
	visioFontFamilies,
} from './ui/formatting';
import { cell, fixture, shape, rectangle, section } from './test-fixtures';

const fonts =
	'<FaceNames><FaceName ID="0" Name="Arial"/><FaceName ID="1" Name="Calibri"/></FaceNames>';
const text = '<Text>Hello\nWorld</Text>';
const character = (cells = '') => section('Character', `<Row IX="0">${cells}</Row>`);
const paragraph = (cells = '') => section('Paragraph', `<Row IX="0">${cells}</Row>`);
const content = cell('Width', 4) + cell('Height', 2) + rectangle;
const source = (extra = '', attributes = '', document = fonts) =>
	fixture({
		document,
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1', content + extra + text, attributes)}</Shapes>` },
		],
		edit: (zip) => zip.file('custom/preserved.bin', new Uint8Array([0, 1, 255, 87])),
	});
const target = { pageId: '0', shapeId: '1' };
const formatText = { ...target, type: 'format-text' as const };
const formatShape = { ...target, type: 'format-shape' as const };
const read = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;

describe('source-backed Visio formatting', () => {
	it('round-trips uniform font family, size, style bits and alignment, preserving all other parts', async () => {
		const bytes = await source(character(cell('Style', 8)) + paragraph(cell('HorzAlign', 0)));
		const original = bytes.slice();
		const saved = await editVsdx(bytes, [
			{
				...formatText,
				fontFamily: 'Calibri',
				fontSize: 18,
				bold: true,
				italic: true,
				underline: true,
				horizontalAlign: 'right',
				verticalAlign: 'top',
			},
		]);
		expect(bytes).toEqual(original);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(saved.diagnostics.map((note) => note.code)).toEqual(['edit-formatting-experimental']);
		const model = await parseVsdx(saved.bytes);
		const shape = model.pages[0]!.shapes[0]!;
		expect(shape.text).toMatchObject({
			plainText: 'Hello\nWorld',
			fontFamily: 'Calibri',
			fontSize: 18 / 72,
			horizontalAlign: 'right',
			verticalAlign: 'top',
		});
		expect(shape.text.runs[0]).toMatchObject({ bold: true, italic: true, underline: true });
		expect(shape.text.paragraphs?.map((p) => p.horizontalAlign)).toEqual(['right', 'right']);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved.bytes);
		for (const path of Object.keys(before.files)) {
			if (before.files[path]!.dir || path === 'visio/pages/page1.xml') continue;
			expect(await after.file(path)!.async('uint8array')).toEqual(
				await before.file(path)!.async('uint8array'),
			);
		}
		expect(await after.file('visio/pages/page1.xml')!.async('string')).toContain(
			'N="Style" V="15"',
		);
		expect(visioFontFamilies(model)).toEqual(['Arial', 'Calibri']);
		expect(visioFontFamilies(copySnapshotScene(model))).toEqual(['Arial', 'Calibri']);
		expect(visioFormattingShape(model.pages[0]!, '1')).toBe(shape);
	});

	it('round-trips opaque solid fill and line color and line point weight', async () => {
		const command = {
			...formatShape,
			fillColor: '#AbCdEf',
			lineColor: '#123456',
			lineWeight: 2.25,
		};
		const saved = await editVsdx(
			await source(cell('FillForegndTrans', 0.5) + cell('LineColorTrans', 0.25)),
			[command],
		);
		expect((await read(saved.bytes)).style).toMatchObject({
			fill: '#abcdef',
			fillOpacity: 1,
			lineColor: '#123456',
			lineOpacity: 1,
			lineWidth: 2.25 / 72,
		});
		const xml = await (
			await JSZip.loadAsync(saved.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		expect(xml).toContain('F="RGB(171,205,239)"');
		const unchanged = await editVsdx(saved.bytes, [command]);
		expect(unchanged.changedParts).toEqual([]);
		expect(unchanged.bytes).toEqual(saved.bytes);
	});

	it('clears native gradient flags while preserving gradient definitions', async () => {
		const gradient = section(
			'FillGradient',
			`<Row IX="0">${cell('GradientStopColor', '#ff0000')}</Row>`,
		);
		const saved = await editVsdx(await source(cell('FillGradientEnabled', 1) + gradient), [
			{ ...formatShape, fillColor: '#ffeedd' },
		]);
		expect((await read(saved.bytes)).style.fillGradient).toBeUndefined();
		const xml = await (
			await JSZip.loadAsync(saved.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		expect(xml).toContain('N="FillGradientEnabled" V="0"');
		expect(xml).toContain('N="GradientStopColor" V="#ff0000"');
	});

	it('supports no fill and zero line weight without guessing geometry', async () => {
		const saved = await editVsdx(await source(), [
			{ ...formatShape, fillColor: 'none', lineWeight: 0 },
		]);
		expect((await read(saved.bytes)).style).toMatchObject({ fill: 'none', lineWidth: 0 });
	});

	it('can format uniform native zero-index markers while preserving text bytes', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', content + character() + '<Text><cp IX="0"/>Hello<pp IX="0"/> world</Text>')}</Shapes>`,
				},
			],
		});
		const saved = await editVsdx(bytes, [{ ...formatText, bold: true }]);
		expect((await read(saved.bytes)).text.plainText).toBe('Hello world');
		const xml = await (
			await JSZip.loadAsync(saved.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		expect(xml).toContain('<Text><cp IX="0"/>Hello<pp IX="0"/> world</Text>');
	});

	it('uses inherited style bits while retaining non-target style definitions', async () => {
		const document =
			fonts +
			`<StyleSheets><StyleSheet ID="2">${character(cell('Style', 8))}</StyleSheet></StyleSheets>`;
		const saved = await editVsdx(await source('', 'TextStyle="2"', document), [
			{ ...formatText, bold: true },
		]);
		const xml = await (
			await JSZip.loadAsync(saved.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		expect(xml).toContain('N="Style" V="9"');
	});
	it('ignores category-disabled style bits instead of reintroducing their formatting', async () => {
		const document = `<StyleSheets><StyleSheet ID="2">${cell('EnableTextProps', 0)}${character(cell('Style', 2))}</StyleSheet></StyleSheets>`;
		const bytes = await source('', 'TextStyle="2"', document);
		expect((await read(bytes)).text.runs[0]!.italic).toBe(false);
		const saved = await editVsdx(bytes, [{ ...formatText, bold: true }]);
		expect((await read(saved.bytes)).text.runs[0]).toMatchObject({ bold: true, italic: false });
	});

	it('snapshots formatting patches before awaiting and strips host properties', async () => {
		const bytes = await source();
		const command = { ...formatText, fontSize: 18, extra: 'host state' };
		const result = editVsdx(bytes, [command]);
		command.fontSize = 300;
		expect((await read((await result).bytes)).text.fontSize).toBe(18 / 72);
		expect(snapshotEdits([command])).toEqual([{ ...formatText, fontSize: 300 }]);
	});

	it('keeps a failed multi-command transaction atomic', async () => {
		const bytes = await source(character(cell('Size', 0.25, 'GUARD(0.25)')));
		const before = bytes.slice();
		await expect(
			editVsdx(bytes, [
				{ ...formatShape, fillColor: '#ff0000' },
				{ ...formatText, fontSize: 20 },
			]),
		).rejects.toThrow(/GUARD|protected|overwrite/i);
		expect(bytes).toEqual(before);
	});
});

describe('formatting admission', () => {
	it.each([
		{ ...formatText },
		{ ...formatShape },
		{ ...formatText, fontSize: 0 },
		{ ...formatText, fontSize: NaN },
		{ ...formatText, fontSize: 1001 },
		{ ...formatText, fontFamily: '' },
		{ ...formatText, fontFamily: 'Bad\nfont' },
		{ ...formatText, bold: 1 },
		{ ...formatText, horizontalAlign: 'distributed' },
		{ ...formatText, verticalAlign: 'left' },
		{ ...formatShape, fillColor: '#fff' },
		{ ...formatShape, lineColor: 'none' },
		{ ...formatShape, lineWeight: -1 },
		{ ...formatShape, lineWeight: Infinity },
	])('rejects invalid patch %#', async (command) => {
		await expect(editVsdx(await source(), [command as VisioFormatEdit])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
	});

	it.each(['Master="2"', 'MasterShape="1"', 'Type="Group"', 'Del="0"'])(
		'rejects unsupported shape metadata %s',
		async (attrs) => {
			await expect(
				editVsdx(await source('', attrs), [{ ...formatText, bold: true }]),
			).rejects.toThrow(/Master|group|deleted/i);
		},
	);
	it.each(['GUARD(0.25)', 'SETATREF(Width)', 'Inh', 'Width/10'])(
		'preserves protected Size formula %s',
		async (formula) => {
			await expect(
				editVsdx(await source(character(cell('Size', 0.25, formula))), [
					{ ...formatText, fontSize: 20 },
				]),
			).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
		},
	);
	it.each(['LockFormat', 'LockTextEdit'])('enforces local %s', async (lock) => {
		await expect(
			editVsdx(await source(cell(lock, 1)), [{ ...formatText, bold: true }]),
		).rejects.toThrow(/protection/i);
	});
	it('checks inherited locks even when local cache overrides the lock', async () => {
		const document = `<StyleSheets><StyleSheet ID="0">${cell('LockFormat', 1)}</StyleSheet></StyleSheets>`;
		await expect(
			editVsdx(await source(cell('LockFormat', 0), '', document), [
				{ ...formatShape, fillColor: '#ff0000' },
			]),
		).rejects.toThrow(/protection/i);
	});
	it('checks lock formula caches and units', async () => {
		await expect(
			editVsdx(await source(cell('LockFormat', 0, '1')), [{ ...formatShape, lineWeight: 2 }]),
		).rejects.toThrow(/stale/i);
		await expect(
			editVsdx(await source('<Cell N="LockFormat" V="0" U="IN"/>'), [
				{ ...formatShape, lineWeight: 2 },
			]),
		).rejects.toThrow(/protection/i);
	});
	it('rejects missing family tables and unknown families', async () => {
		await expect(
			editVsdx(await source('', '', ''), [{ ...formatText, fontFamily: 'Arial' }]),
		).rejects.toThrow(/FaceNames/i);
		await expect(
			editVsdx(await source(), [{ ...formatText, fontFamily: 'Invented' }]),
		).rejects.toThrow(/FaceName/i);
	});
	it('rejects duplicate and deleted formatting rows', async () => {
		for (const extra of [
			section('Character', '<Row IX="0"/><Row IX="0"/>'),
			section('Character', '<Row IX="0" Del="1"/>'),
			section('Character', '<Row IX="1" Del="1"/>'),
		])
			await expect(editVsdx(await source(extra), [{ ...formatText, bold: true }])).rejects.toThrow(
				/format|fields/i,
			);
	});
	it('rejects error caches, invalid length units and ambiguous cell names', async () => {
		for (const extra of [
			character('<Cell N="Size" V="0.25" E="error"/>'),
			character('<Cell N="Size" V="0.25" U="DEG"/>'),
			character(cell('Size', 0.25) + cell('size', 0.25)),
			character(cell('size', 0.25)),
		])
			await expect(
				editVsdx(await source(extra), [{ ...formatText, fontSize: 20 }]),
			).rejects.toThrow(/error|unit|unique|noncanonical/i);
	});
	it.each(['Character.0.Size', 'Char.Size'])(
		'rejects dependent page formulas for %s',
		async (name) => {
			await expect(
				editVsdx(await source(cell('UserCache', 0.5, `Sheet.1!${name}*2`)), [
					{ ...formatText, fontSize: 20 },
				]),
			).rejects.toThrow(/dependent caches/i);
		},
	);
	it('rejects dependent package metadata and dynamic expressions', async () => {
		await expect(
			editVsdx(await source('', '', cell('UserCache', 1, 'Sheet.1!LineWeight')), [
				{ ...formatShape, lineWeight: 2 },
			]),
		).rejects.toThrow(/dependent caches/i);
		await expect(
			editVsdx(await source(cell('UserCache', 1, 'INDIRECT(&quot;Width&quot;)')), [
				{ ...formatShape, lineWeight: 2 },
			]),
		).rejects.toThrow(/dynamic|dependencies/i);
	});
	it('validates host font list shape, names and aggregate metadata limits', async () => {
		const model = await parseVsdx(await source());
		for (const families of [
			'Arial',
			[1],
			['x'.repeat(1025)],
			Array(10001).fill('Arial'),
			Array(1000).fill('x'.repeat(1024)),
		]) {
			expect(() =>
				assertViewableDocument({ ...model, fontFamilies: families } as typeof model),
			).toThrow(/font family|metadata/i);
		}
	});
	it('allows shape paint controls independently of mixed text styles', async () => {
		const model = await parseVsdx(await source());
		const page = model.pages[0]!,
			shape = page.shapes[0]!;
		shape.text.runs.push({ ...shape.text.runs[0]!, bold: true });
		expect(visioFormattingShape(page, '1')).toBe(shape);
		expect(visioStyleFormattingShape(page, '1')).toBe(shape);
	});
	it('rejects local and inherited layer membership conservatively', async () => {
		await expect(
			editVsdx(await source(cell('LayerMember', 0)), [{ ...formatShape, fillColor: '#ff0000' }]),
		).rejects.toThrow(/layered/i);
		const document = `<StyleSheets><StyleSheet ID="0">${cell('LayerMember', 0)}</StyleSheet></StyleSheets>`;
		await expect(
			editVsdx(await source('', '', document), [{ ...formatText, bold: true }]),
		).rejects.toThrow(/layered/i);
	});
});

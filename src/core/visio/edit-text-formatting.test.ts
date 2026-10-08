import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx, type VisioTextFormatEdit } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, shape, rectangle, section } from './test-fixtures';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';
import { visioTextFormattingState, visioTextIndentCommand } from './ui/formatting';

const command = { type: 'format-text' as const, pageId: '0', shapeId: '1' };
const row = (name: string, cells = '') => section(name, `<Row IX="0">${cells}</Row>`);
const content = (extra = '', text = 'First words\nSecond words') =>
	cell('Width', 4) + cell('Height', 2) + rectangle + extra + `<Text>${text}</Text>`;
const source = (extra = '', document = '') =>
	fixture({
		document,
		pages: [{ id: '0', contents: `<Shapes>${shape('1', content(extra))}</Shapes>` }],
		edit: (zip) => zip.file('unknown/binary.dat', new Uint8Array([0, 42, 255])),
	});
const parsed = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!.text;
const xml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');

describe('extended whole-shape text formatting', () => {
	it('round-trips opaque font color, separate strikethrough, bullets, indent and justified paragraphs', async () => {
		const bytes = await source(
			row('Character', cell('Style', 5) + cell('ColorTrans', 0.5)) +
				row(
					'Paragraph',
					cell('SpLine', -1.5) + cell('IndRight', 0.2) + cell('TextPosAfterBullet', 0.1),
				),
		);
		const patch = {
			...command,
			fontColor: '#123ABC',
			strikethrough: true,
			bullets: true,
			indentLeft: 18,
			horizontalAlign: 'justify' as const,
		};
		const saved = await editVsdx(bytes, [patch]);
		const text = await parsed(saved.bytes);
		expect(text).toMatchObject({
			color: '#123abc',
			horizontalAlign: 'justify',
			strikethrough: true,
			bold: true,
			underline: true,
		});
		expect(text.opacity).toBeUndefined();
		expect(text.runs.every((run) => run.strikethrough && run.bold && run.underline)).toBe(true);
		expect(
			text.paragraphs?.every(
				(paragraph) =>
					paragraph.horizontalAlign === 'justify' &&
					paragraph.indentLeft === 0.25 &&
					paragraph.bullet?.text === '•',
			),
		).toBe(true);
		expect(text.paragraphs?.[0]).toMatchObject({
			indentRight: 0.2,
			lineSpacing: { value: 1.5 },
			bullet: { offset: 0.1 },
		});
		const savedXml = await xml(saved.bytes);
		expect(savedXml).toContain('N="Color" V="#123abc" F="RGB(18,58,188)"');
		expect(savedXml).toContain('N="Strikethru" V="1"');
		expect(savedXml).toContain('N="IndLeft" V="0.25" U="PT"');
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved.bytes);
		for (const path of Object.keys(before.files))
			if (!before.files[path]!.dir && path !== 'visio/pages/page1.xml')
				expect(await after.file(path)!.async('uint8array')).toEqual(
					await before.file(path)!.async('uint8array'),
				);
		const unchanged = await editVsdx(saved.bytes, [patch]);
		expect(unchanged.changedParts).toEqual([]);
		expect(unchanged.bytes).toEqual(saved.bytes);
	});

	it('toggles strikethrough without changing underline, double strike or text payloads', async () => {
		const bytes = await source(
			row('Character', cell('Style', 4) + cell('DoubleStrikethrough', 1) + cell('Strikethru', 1)),
			'<Custom><Unrecognized>retain</Unrecognized></Custom>',
		);
		const saved = await editVsdx(bytes, [{ ...command, strikethrough: false }]);
		expect((await parsed(saved.bytes)).runs[0]).toMatchObject({
			underline: true,
			strikethrough: false,
		});
		expect(await xml(saved.bytes)).toContain('N="DoubleStrikethrough" V="1"');
	});

	it('keeps custom and builtin bullet glyph selections when enabling bullets', async () => {
		const builtin = await editVsdx(await source(row('Paragraph', cell('Bullet', 2))), [
			{ ...command, bullets: true },
		]);
		expect(builtin.changedParts).toEqual([]);
		const custom = await editVsdx(
			await source(
				row('Paragraph', cell('Bullet', 0) + '<Cell N="BulletStr" V="ooo" F="&quot;ooo&quot;"/>'),
			),
			[{ ...command, bullets: true }],
		);
		expect((await parsed(custom.bytes)).paragraphs?.[0]?.bullet?.text).toBe('ooo');
		const without = await editVsdx(custom.bytes, [{ ...command, bullets: false }]);
		expect((await parsed(without.bytes)).paragraphs?.[0]?.bullet).toBeUndefined();
		expect(await xml(without.bytes)).toContain('N="BulletStr" V="ooo" F="&quot;ooo&quot;"');
	});

	it('uses inherited paragraph values and category gating without changing parent style definitions', async () => {
		const document = `<StyleSheets><StyleSheet ID="2">${row('Paragraph', cell('Bullet', 3) + cell('IndLeft', 0.5))}</StyleSheet></StyleSheets>`;
		const bytes = await fixture({
			document,
			pages: [{ id: '0', contents: `<Shapes>${shape('1', content(), 'TextStyle="2"')}</Shapes>` }],
		});
		const saved = await editVsdx(bytes, [{ ...command, bullets: true, indentLeft: 54 }]);
		expect((await parsed(saved.bytes)).paragraphs?.[0]).toMatchObject({
			indentLeft: 0.75,
			bullet: { text: '■' },
		});
	});

	it('keeps malformed rows and unknown markup untouched when refusing an unsupported target', async () => {
		const bytes = await source(
			section('Character', `<Row IX="0"/><Row IX="01">${cell('Style', 1)}</Row>`),
		);
		const before = bytes.slice();
		await expect(editVsdx(bytes, [{ ...command, fontColor: '#112233' }])).rejects.toThrow(/row/i);
		expect(bytes).toEqual(before);
		const guarded = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', content('', 'Hello<future Custom="retained"/>'))}</Shapes>`,
				},
			],
		});
		await expect(editVsdx(guarded, [{ ...command, strikethrough: true }])).rejects.toThrow(
			/unknown text markup/i,
		);
		expect(await xml(guarded)).toContain('future Custom="retained"');
	});

	it('refuses a guarded later property atomically after an earlier color edit', async () => {
		const bytes = await source(row('Paragraph', cell('IndLeft', 0, 'GUARD(0)')));
		const before = bytes.slice();
		await expect(
			editVsdx(bytes, [
				{ ...command, fontColor: '#112233' },
				{ ...command, indentLeft: 18 },
			]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
		expect(bytes).toEqual(before);
	});
	it('admits native empty LayerMember caches and preserves their value', async () => {
		const saved = await editVsdx(await source(cell('LayerMember', '')), [
			{ ...command, fontColor: '#112233' },
		]);
		expect((await parsed(saved.bytes)).color).toBe('#112233');
		expect(await xml(saved.bytes)).toContain('N="LayerMember" V=""');
	});
	it.each([
		cell('LayerMember', '', '&quot;&quot;'),
		'<Cell N="LayerMember" V="" E="error"/>',
		'<Cell N="LayerMember"/>',
		cell('LayerMember', '', 'Inh'),
	])('rejects unresolved empty layer membership %#', async (membership) => {
		await expect(
			editVsdx(await source(membership), [{ ...command, fontColor: '#112233' }]),
		).rejects.toThrow(/layer|inherited/i);
	});

	it.each(['Char.Color', 'Char.Strikethru', 'Para.Bullet', 'Para.IndLeft', 'Para.HorzAlign'])(
		'refuses stale dependent formula cache for %s',
		async (name) => {
			await expect(
				editVsdx(await source(cell('UserCache', 1, `${name}*2`)), [
					{
						...command,
						fontColor: '#112233',
						strikethrough: true,
						bullets: true,
						indentLeft: 18,
						horizontalAlign: 'justify',
					},
				]),
			).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_FORMAT_DEPENDENCY' });
		},
	);

	it.each([
		{ fontColor: '#fff' },
		{ fontColor: 'none' },
		{ strikethrough: 1 },
		{ bullets: 'round' },
		{ indentLeft: NaN },
		{ indentLeft: Infinity },
		{ indentLeft: -1 },
		{ indentLeft: 7201 },
	])('validates new patch %#', async (patch) => {
		await expect(
			editVsdx(await source(), [{ ...command, ...patch } as VisioTextFormatEdit]),
		).rejects.toMatchObject({ code: 'INVALID_EDIT' });
	});

	it('snapshots new fields before asynchronous work and does not scale physical indent points', async () => {
		const bytes = await source();
		const patch = {
			...command,
			fontColor: '#112233',
			strikethrough: true,
			bullets: true,
			indentLeft: 18,
			horizontalAlign: 'justify' as const,
		};
		const pending = editVsdx(bytes, [patch]);
		patch.fontColor = '#ffffff';
		patch.indentLeft = 144;
		const text = await parsed((await pending).bytes);
		expect(text.color).toBe('#112233');
		expect(text.paragraphs?.[0]?.indentLeft).toBe(0.25);
		expect(snapshotEdits([patch])[0]).toEqual(patch);
		const page = (await parseVsdx(bytes)).pages[0]!;
		expect(visioPageEditToDrawing({ ...page, drawingToPageScale: 0.5 }, patch)).toEqual(patch);
	});
});

describe('text formatting scene contracts', () => {
	it('preserves default and run strikethrough through bounded scene snapshots', async () => {
		const saved = await editVsdx(await source(), [
			{ ...command, strikethrough: true, horizontalAlign: 'justify' },
		]);
		const model = await parseVsdx(saved.bytes);
		assertViewableDocument(model);
		const copy = copySnapshotScene(model);
		expect(copy.pages[0]!.shapes[0]!.text).toMatchObject({
			strikethrough: true,
			horizontalAlign: 'justify',
		});
		expect(copy.pages[0]!.shapes[0]!.text.runs[0]?.strikethrough).toBe(true);
		copy.pages[0]!.shapes[0]!.text.runs[0]!.strikethrough = 'invalid' as unknown as boolean;
		expect(() => assertViewableDocument(copy)).toThrow(/strikethrough/i);
	});

	it('aggregates mixed selection without inventing common font or toggled states', async () => {
		const model = await parseVsdx(await source());
		const a = model.pages[0]!.shapes[0]!;
		const b = structuredClone(a);
		b.id = '2';
		a.text.runs[0]!.bold = true;
		b.text.runs[0]!.fontSize *= 2;
		b.text.runs[0]!.color = '#112233';
		b.text.paragraphs![0]!.indentLeft = 0.25;
		expect(visioTextFormattingState([a, b])).toMatchObject({
			bold: false,
			italic: false,
			strikethrough: false,
			bullets: false,
			fontFamily: 'Arial',
			fontSize: undefined,
			fontColor: undefined,
			canIndentDecrease: true,
		});
		expect(visioTextFormattingState([])).toMatchObject({
			bold: false,
			bullets: false,
			fontFamily: undefined,
			canIndentDecrease: false,
		});
	});

	it('returns bounded absolute indent commands and refuses inconsistent paragraph offsets', async () => {
		const model = await parseVsdx(await source());
		const page = model.pages[0]!;
		expect(visioTextIndentCommand(page, '1', 'increase')).toEqual({ ...command, indentLeft: 18 });
		expect(visioTextIndentCommand(page, '1', 'decrease')).toEqual({ ...command, indentLeft: 0 });
		page.shapes[0]!.text.paragraphs![0]!.indentLeft = 0.25;
		expect(visioTextIndentCommand(page, '1', 'increase')).toBeUndefined();
	});
});

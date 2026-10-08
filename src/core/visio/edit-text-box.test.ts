import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { attribute, children } from './sheet';
import { cell, fixture, ns, rectangle, shape } from './test-fixtures';
import { snapshotVisioEdits, type VisioEdit } from './edit-commands';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';

const create = {
	type: 'create-text-box' as const,
	pageId: '0',
	shapeId: '2',
	x: 3,
	y: 4,
	width: 2,
	height: 1,
	text: 'Text & <label> 😀\n',
};
const replace = { type: 'replace-plain-text' as const, pageId: '0', shapeId: '1', text: 'New' };
const styles = (lock = '', extra = '') =>
	`<DocumentSettings DefaultLineStyle="3" DefaultFillStyle="3" DefaultTextStyle="3"/><StyleSheets><StyleSheet ID="3">${lock}${cell('LeftMargin', 0.1)}${cell('VerticalAlign', 0)}<Section N="Character"><Row IX="0">${cell('Size', 0.25)}</Row></Section><Section N="Paragraph"><Row IX="0">${cell('HorzAlign', 2)}</Row></Section>${extra}</StyleSheet></StyleSheets>`;
const source = (document = '', contents = '') =>
	fixture({
		document,
		pages: [
			{ id: '0', contents: contents || `<Shapes>${shape('1', '<Text>Old</Text>')}</Shapes>` },
		],
		edit: (zip) => zip.file('custom/opaque.bin', new Uint8Array([0, 128, 255])),
	});

describe('source-backed fixed text box creation', () => {
	it('preserves default text styles, logical trailing newlines and untouched payloads', async () => {
		const bytes = await source(styles());
		const saved = await editVsdx(bytes, [create]);
		const actual = (await parseVsdx(saved.bytes)).pages[0]!.shapes[1]!;
		expect(actual).toMatchObject({ width: 2, height: 1, rotation: { pinX: 3, pinY: 4 } });
		expect(actual.style).toMatchObject({ fill: 'none', linePattern: 0 });
		expect(actual.text).toMatchObject({
			plainText: create.text,
			fontSize: 0.25,
			horizontalAlign: 'right',
			verticalAlign: 'top',
			margins: { left: 0.1 },
		});
		const before = await VisioPackage.open(bytes),
			after = await VisioPackage.open(saved.bytes);
		for (const path of before.paths())
			if (path !== 'visio/pages/page1.xml')
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		const root = await after.readXml('visio/pages/page1.xml');
		const created = children(children(root, 'Shapes')[0], 'Shape')[1]!;
		expect(['LineStyle', 'FillStyle', 'TextStyle'].map((name) => attribute(created, name))).toEqual(
			['3', '3', '3'],
		);
		expect(children(created, 'Text')[0]!.textContent).toBe(create.text + '\n');
		expect(children(created, 'Section')[0]!.namespaceURI).toBe(ns);
	});
	it('creates the Shapes container before existing Connects on an empty page', async () => {
		const saved = await editVsdx(await source('', '<Connects/>'), [create]);
		const root = await (await VisioPackage.open(saved.bytes)).readXml('visio/pages/page1.xml');
		expect(
			Array.from(root.childNodes)
				.filter((n) => n.nodeType === 1)
				.map((n) => (n as Element).localName),
		).toEqual(['Shapes', 'Connects']);
		expect(children(root, 'Shapes')[0]!.namespaceURI).toBe(ns);
	});
	it('remains movable, resizable and eligible for later fill and line color formatting', async () => {
		const bytes = (await editVsdx(await source(), [create])).bytes;
		const saved = await editVsdx(bytes, [
			{ type: 'move-shape', pageId: '0', shapeId: '2', x: 5, y: 6 },
			{
				type: 'resize-shape',
				pageId: '0',
				shapeId: '2',
				width: 4,
				height: 2,
				anchor: { x: 0, y: 0 },
			},
			{
				type: 'format-shape',
				pageId: '0',
				shapeId: '2',
				fillColor: '#abcdef',
				lineColor: '#123456',
			},
		]);
		const target = (await parseVsdx(saved.bytes)).pages[0]!.shapes[1]!;
		expect(target).toMatchObject({ width: 4, height: 2, text: { plainText: create.text } });
		expect(target.style).toMatchObject({ fill: '#abcdef', lineColor: '#123456', linePattern: 0 });
		expect(target.geometry.every((path) => path.fill && path.stroke)).toBe(true);
	});
	it.each(['', '\n', '\n\n', 'A\n', 'A\n\n', '  spaced\t😀  '])(
		'roundtrips exact caller text %j',
		async (text) => {
			const saved = await editVsdx(await source(), [{ ...create, text }]);
			expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[1]!.text.plainText).toBe(text);
			const unchanged = await editVsdx(saved.bytes, [{ ...replace, shapeId: '2', text }]);
			expect(unchanged.bytes).toEqual(saved.bytes);
		},
	);
	it('owns source and creation text before asynchronous processing', async () => {
		const bytes = await source(),
			command = { ...create };
		const pending = editVsdx(bytes, [command]);
		bytes.fill(0);
		command.text = 'Changed';
		command.width = 9;
		expect((await parseVsdx((await pending).bytes)).pages[0]!.shapes[1]!).toMatchObject({
			width: 2,
			text: { plainText: create.text },
		});
	});
	it('ignores inactive inherited paint formulas that are explicitly overridden', async () => {
		const saved = await editVsdx(
			await source(styles('', cell('FillPattern', 1, 'Width') + cell('LinePattern', 1, 'Height'))),
			[create],
		);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[1]!.style).toMatchObject({
			fill: 'none',
			linePattern: 0,
		});
	});
	it.each([
		`<Shapes>${shape('1', `<Shapes>${shape('2', '<Text>Nested</Text>')}</Shapes>`)}</Shapes>`,
		`<Shapes>${shape('1', `<Unknown>${shape('2', '<Text>Concealed</Text>')}</Unknown>`)}</Shapes>`,
		`<Shapes>${shape('1')}${shape('1')}</Shapes>`,
	])('rejects source identity collisions or ambiguity atomically', async (contents) => {
		const bytes = await source('', contents),
			copy = bytes.slice();
		await expect(editVsdx(bytes, [create])).rejects.toMatchObject({ code: 'INVALID_SHAPE_ID' });
		expect(bytes).toEqual(copy);
	});
	it.each(['Sheet.2!Width', 'SHAPETEXT(Sheet.2!TheText)', 'CONTAINERSHEETREF(1)'])(
		'refuses prospective ID or container dependencies: %s',
		async (formula) => {
			const bytes = await source('', `<Shapes>${shape('1', cell('PinX', 1, formula))}</Shapes>`);
			await expect(editVsdx(bytes, [create])).rejects.toThrow();
		},
	);
	it('refuses a mixed transaction atomically when the later text creation is protected', async () => {
		const bytes = await source(styles(cell('LockTextEdit', 1))),
			copy = bytes.slice();
		await expect(
			editVsdx(bytes, [
				{ type: 'create-rectangle', pageId: '0', shapeId: '3', x: 2, y: 2, width: 1, height: 1 },
				create,
			]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
		expect(bytes).toEqual(copy);
	});
	it('shares page-to-drawing scaling and strips arbitrary command properties', async () => {
		const page = (
			await parseVsdx(
				await fixture({
					pages: [
						{ id: '0', contents: '', pageCells: cell('PageScale', 1) + cell('DrawingScale', 2) },
					],
				}),
			)
		).pages[0]!;
		expect(visioPageEditToDrawing(page, create)).toMatchObject({
			x: 6,
			y: 8,
			width: 4,
			height: 2,
			text: create.text,
		});
		const arbitrary = { ...create, hostile: { data: 1 } };
		for (const snapshot of [snapshotEdits([arbitrary]), snapshotVisioEdits([arbitrary], 1000, 1e6)])
			expect(snapshot).toEqual([create]);
	});
	it.each([
		{ text: undefined },
		{ width: 0 },
		{ height: -1 },
		{ x: Infinity },
		{ shapeId: '02' },
		{ text: '\r' },
		{ text: '\ud800' },
	])('rejects invalid creation commands: %j', async (patch) => {
		await expect(
			editVsdx(await source(), [{ ...create, ...patch } as VisioEdit]),
		).rejects.toThrow();
	});
	it('counts required creation text in the aggregate command limit', async () => {
		await expect(
			editVsdx(await source(), [create], { maxTextCharacters: 2 }),
		).rejects.toMatchObject({ code: 'LIMIT_EDIT_TEXT' });
	});
});

describe('effective plain-text edit protection', () => {
	const locks = [
		cell('LockTextEdit', 1),
		cell('LockTextEdit', 0, '1'),
		cell('LockTextEdit', 0, 'UNKNOWN(0)'),
		cell('LockTextEdit', 0, 'Width*0'),
		cell('LockTextEdit', 0, 'SETATREF(User.flag)'),
		cell('locktextedit', 1),
		'<Cell N="LockTextEdit" V="0" U="RAD"/>',
		'<Cell N="LockTextEdit" V="0" E="#VALUE!"/>',
	];
	it.each(locks)('refuses protected, stale, erroneous or ambiguous local lock %s', async (lock) => {
		const bytes = await source('', `<Shapes>${shape('1', lock + '<Text>Old</Text>')}</Shapes>`),
			copy = bytes.slice();
		await expect(editVsdx(bytes, [replace])).rejects.toThrow();
		expect(bytes).toEqual(copy);
	});
	it.each(locks)('refuses the same inherited creation lock %s', async (lock) => {
		await expect(editVsdx(await source(styles(lock)), [create])).rejects.toThrow();
	});
	it('allows a proven inactive guarded lock without rewriting it or unrelated geometry', async () => {
		const bytes = await source(
			'',
			`<Shapes>${shape('1', cell('LockTextEdit', 0, 'GUARD(0)') + rectangle + '<Text>Old</Text>')}</Shapes>`,
		);
		const saved = await editVsdx(bytes, [replace]);
		const root = await (await VisioPackage.open(saved.bytes)).readXml('visio/pages/page1.xml');
		const target = children(children(root, 'Shapes')[0], 'Shape')[0]!;
		expect(attribute(children(target, 'Cell')[0], 'F')).toBe('GUARD(0)');
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.text.plainText).toBe('New');
	});
	it('does not apply an unrelated sibling lock to an unlocked text target', async () => {
		const bytes = await source(
			'',
			`<Shapes>${shape('1', '<Text>Old</Text>')}${shape('3', cell('LockTextEdit', 1) + '<Text>Locked</Text>')}</Shapes>`,
		);
		expect(
			(await parseVsdx((await editVsdx(bytes, [replace])).bytes)).pages[0]!.shapes[0]!.text
				.plainText,
		).toBe('New');
	});
	it('matches native target-only text locks for a child of a locked group', async () => {
		const bytes = await source(
			'',
			`<Shapes>${shape('1', cell('LockTextEdit', 1) + `<Text>Group</Text><Shapes>${shape('2', '<Text>Child</Text>')}</Shapes>`, 'Type="Group"')}</Shapes>`,
		);
		const saved = await editVsdx(bytes, [{ ...replace, shapeId: '2' }]);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.children[0]!.text.plainText).toBe(
			'New',
		);
		await expect(editVsdx(bytes, [replace])).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
	});
	it('resolves explicit style ancestry conservatively for existing plain replacement', async () => {
		const document = `<StyleSheets><StyleSheet ID="3" TextStyle="4"/><StyleSheet ID="4">${cell('LockTextEdit', 1)}</StyleSheet></StyleSheets>`;
		const bytes = await source(
			document,
			`<Shapes>${shape('1', '<Text>Old</Text>', 'TextStyle="3"')}</Shapes>`,
		);
		await expect(editVsdx(bytes, [replace])).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
	});
});

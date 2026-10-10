import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, shape } from './test-fixtures';

const PAGE = 'visio/pages/page1.xml';
const MASTER = 'visio/masters/master1.xml';
const size = cell('Width', 2) + cell('Height', 1, 'User.ResizeTxtHeight');
/** A stencil master whose height follows its text, as Visio's flowchart masters do. */
const master = (extra = '') =>
	shape(
		'5',
		size +
			`<Section N="User"><Row N="ResizeTxtHeight">${cell('Value', 1, 'MAX(1,TEXTHEIGHT(TheText,TxtWidth))')}</Row></Section>` +
			rectangle +
			extra,
	);
const instance = (contents = '') =>
	shape('1', cell('PinX', 2) + cell('PinY', 6) + contents, 'Type="Shape" Master="2"');
const source = (page: string, definition = master()) =>
	fixture({
		masters: [{ id: '2', shapes: definition }],
		pages: [{ id: '0', contents: `<Shapes>${page}</Shapes>` }],
	});
const part = async (bytes: Uint8Array, path: string) =>
	(await JSZip.loadAsync(bytes)).file(path)!.async('string');
const type = (bytes: Uint8Array, text: string, shapeId = '1') =>
	editVsdx(bytes, [{ type: 'replace-plain-text', pageId: '0', shapeId, text }]);

describe('text of stencil (master) instances', () => {
	it('writes a local Text element and nothing else, as Visio does', async () => {
		const bytes = await source(instance());
		const saved = await type(bytes, 'Typed on an instance');
		expect(saved.changedParts).toEqual([PAGE]);
		const page = await part(saved.bytes, PAGE);
		// The instance keeps its cells and gains the text after them; no size cell is written.
		expect(page).toMatch(
			/<Cell N="PinY" V="6"\/><Text>Typed on an instance\n<\/Text><\/Shape>/,
		);
		expect(page).not.toMatch(/N="Height"/);
		expect(await part(saved.bytes, MASTER)).toBe(await part(bytes, MASTER));
		const model = await parseVsdx(saved.bytes);
		expect(model.pages[0]!.shapes[0]!.text.plainText).toBe('Typed on an instance');
	});

	it('replaces an existing local text and skips unchanged or inherited text', async () => {
		const bytes = await source(instance('<Text>First\n</Text>'));
		const saved = await type(bytes, 'Second');
		const page = await part(saved.bytes, PAGE);
		expect(page.match(/<Text>/g)).toHaveLength(1);
		expect(page).toContain('<Text>Second\n</Text>');
		expect((await type(saved.bytes, 'Second')).changedParts).toEqual([]);
		// Typing the master's own text changes nothing.
		const inherited = await source(instance(), master('<Text>From the master\n</Text>'));
		expect((await type(inherited, 'From the master')).changedParts).toEqual([]);
		expect(await part((await type(inherited, 'Mine')).bytes, PAGE)).toContain('<Text>Mine\n</Text>');
	});

	it('edits a sub-shape of a group instance before its own sub-shapes', async () => {
		const group = shape(
			'5',
			size + `<Shapes>${shape('6', size + rectangle)}</Shapes>`,
			'Type="Group"',
		);
		const page = shape(
			'1',
			cell('PinX', 2) +
				`<Shapes>${shape('3', `<Shapes>${shape('4', '', 'MasterShape="6"')}</Shapes>`, 'MasterShape="6"')}</Shapes>`,
			'Type="Group" Master="2"',
		);
		const saved = await type(await source(page, group), 'Inside the group', '3');
		expect(await part(saved.bytes, PAGE)).toMatch(
			/<Shape ID="3" MasterShape="6"><Text>Inside the group\n<\/Text><Shapes>/,
		);
	});

	it('honours text protection from the master unless the instance overrides it', async () => {
		const locked = master(cell('LockTextEdit', 1));
		await expect(type(await source(instance(), locked), 'No')).rejects.toMatchObject({
			code: 'EDIT_PROTECTED_CELL',
		});
		const unlocked = await type(
			await source(instance(cell('LockTextEdit', 0)), locked),
			'Allowed here',
		);
		expect(await part(unlocked.bytes, PAGE)).toContain('<Text>Allowed here\n</Text>');
		await expect(
			type(await source(instance(cell('LockTextEdit', 1))), 'No'),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
	});

	it('refuses master text with fields or rich markup, and a missing master', async () => {
		const rich = master('<Text><cp IX="0"/>Rich\n</Text>');
		await expect(type(await source(instance(), rich), 'Plain')).rejects.toMatchObject({
			code: 'UNSUPPORTED_TEXT_EDIT',
		});
		const field = master(
			`<Section N="Field">${`<Row IX="0">${cell('Value', 1)}</Row>`}</Section><Text>Page\n</Text>`,
		);
		await expect(type(await source(instance(), field), 'Plain')).rejects.toMatchObject({
			code: 'UNSUPPORTED_TEXT_EDIT',
		});
		const orphan = await fixture({
			masters: [{ id: '9', shapes: master() }],
			pages: [{ id: '0', contents: `<Shapes>${instance()}</Shapes>` }],
		});
		await expect(type(orphan, 'Plain')).rejects.toMatchObject({ code: 'UNSUPPORTED_TEXT_EDIT' });
	});
});

// Optional: `scripts/record-visio-instance-text.ps1 -OutputDirectory <dir>` saves a real stencil
// drop before and after Visio set its text; the edit must write the same Text element.
const native = process.env.VISIO_NATIVE_INSTANCE_TEXT_DIR;
it.skipIf(!native)('writes the Text element native Visio writes for a dropped stencil shape', async () => {
	const before = new Uint8Array(await readFile(join(native!, 'before.vsdx')));
	const after = new Uint8Array(await readFile(join(native!, 'after.vsdx')));
	const saved = await type(before, 'Typed on an instance', '1');
	// An XML reader normalises the file's CRLF to a line feed, so compare the text that way.
	const text = (xml: string) =>
		/<Shape ID=["']1["'][^>]*>.*?<Text>(.*?)<\/Text>/s.exec(xml)?.[1]?.replace(/\r\n/g, '\n');
	expect(text(await part(saved.bytes, PAGE))).toBe(text(await part(after, PAGE)));
	const model = await parseVsdx(saved.bytes);
	const expected = await parseVsdx(after);
	expect(model.pages[0]!.shapes[0]!.text.plainText).toBe(
		expected.pages[0]!.shapes[0]!.text.plainText,
	);
});

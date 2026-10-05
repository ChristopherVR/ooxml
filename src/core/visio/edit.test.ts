import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit.js';
import { parseVsdx } from './parser.js';
import { fixture, shape, cell } from './test-fixtures.js';
const command = (text = 'New & <text> 😀') => ({
	type: 'replace-plain-text' as const,
	pageId: '0',
	shapeId: '1',
	text,
});
const input = (text = '<Text>Old</Text>', attrs = '') =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 3, '1+2') + text, attrs)}</Shapes>`,
			},
		],
	});
describe('experimental VSDX text editing', () => {
	it('saves and reopens plain text with unchanged cached geometry', async () => {
		const original = await input();
		const before = await parseVsdx(original);
		const result = await editVsdx(original, [command()]);
		expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(result.diagnostics[0]?.code).toBe('edit-caches-not-recalculated');
		const after = await parseVsdx(result.bytes);
		expect(after.pages[0]?.shapes[0]?.text.plainText).toBe(command().text);
		expect(after.pages[0]?.shapes[0]?.width).toBe(before.pages[0]?.shapes[0]?.width);
	});
	it('preserves the whole archive for no commands and equal text', async () => {
		const original = await input();
		for (const edits of [[], [command('Old')]]) {
			const result = await editVsdx(original, edits);
			expect(result.bytes).toEqual(original);
			expect(result.bytes).not.toBe(original);
			expect(result.changedParts).toEqual([]);
		}
	});
	it('supports empty text, CDATA and repeated fresh transactions', async () => {
		let bytes = await input('<Text><![CDATA[Before]]></Text>');
		for (const text of ['', 'after', '\tline\nnext', '😀']) {
			bytes = (await editVsdx(bytes, [command(text)])).bytes;
			expect((await parseVsdx(bytes)).pages[0]?.shapes[0]?.text.plainText).toBe(text);
		}
	});
	it('supports nested local shapes with stable page-wide IDs', async () => {
		const bytes = await input(`<Shapes>${shape('2', '<Text>Nested</Text>')}</Shapes>`);
		const result = await editVsdx(bytes, [{ ...command('Changed'), shapeId: '2' }]);
		expect((await parseVsdx(result.bytes)).pages[0]?.shapes[0]?.children[0]?.text.plainText).toBe(
			'Changed',
		);
	});
	it.each([
		'<Text><cp IX="0"/>Old</Text>',
		'<Text><fld IX="0">Old</fld></Text>',
		'<Text><!--note-->Old</Text>',
		'<Text>Old</Text><Text>Other</Text>',
		'',
		'<Text>Old</Text><Section N="Field"/>',
	])('rejects unsupported text without rewriting input: %s', async (text) => {
		const bytes = await input(text),
			copy = bytes.slice();
		await expect(editVsdx(bytes, [command()])).rejects.toMatchObject({
			code: 'UNSUPPORTED_TEXT_EDIT',
		});
		expect(bytes).toEqual(copy);
	});
	it.each(['Master="7"', 'MasterShape="5"', 'Del="1"'])(
		'rejects inherited/deleted targets: %s',
		async (attrs) => {
			await expect(editVsdx(await input(undefined, attrs), [command()])).rejects.toMatchObject({
				code: 'UNSUPPORTED_TEXT_EDIT',
			});
		},
	);
	it('rejects a batch atomically when a later target is absent', async () => {
		const bytes = await input(),
			copy = bytes.slice();
		await expect(
			editVsdx(bytes, [command(), { ...command(), shapeId: 'missing' }]),
		).rejects.toMatchObject({ code: 'EDIT_TARGET_NOT_FOUND' });
		expect(bytes).toEqual(copy);
		expect((await editVsdx(bytes, [])).bytes).toEqual(copy);
	});
	it('owns input and command values before asynchronous processing', async () => {
		const bytes = await input(),
			edit = command('Expected');
		const pending = editVsdx(bytes, [edit]);
		bytes.fill(0);
		edit.text = 'Mutated';
		expect((await parseVsdx((await pending).bytes)).pages[0]?.shapes[0]?.text.plainText).toBe(
			'Expected',
		);
	});
});

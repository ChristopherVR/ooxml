import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import type { VisioShapeHyperlinkEdit } from './edit-metadata-commands';
import { cell, fixture, rectangle, shape } from './test-fixtures';

const link = (patch: Partial<VisioShapeHyperlinkEdit> = {}): VisioShapeHyperlinkEdit => ({
	type: 'set-shape-hyperlink',
	pageId: '0',
	shapeId: '1',
	hyperlink: { address: 'https://example.com/a', subAddress: '', description: 'Example' },
	...patch,
});
async function drawing(): Promise<Uint8Array> {
	return (
		await editVsdx(await createVsdx(), [
			{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 2, width: 2, height: 1 },
		])
	).bytes;
}
const first = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
const page = async (bytes: Uint8Array) =>
	new TextDecoder().decode(
		await (await VisioPackage.open(bytes)).readBytes('visio/pages/page1.xml'),
	);

describe('shape hyperlinks', () => {
	it('adds, edits and removes a local Hyperlink row that reads back through parseVsdx', async () => {
		const added = await editVsdx(await drawing(), [link()]);
		expect(await page(added.bytes)).toContain('<Section N="Hyperlink"><Row N="Row_1">');
		expect((await first(added.bytes)).hyperlinks).toEqual([
			expect.objectContaining({
				name: 'Row_1',
				description: 'Example',
				address: 'https://example.com/a',
				target: { kind: 'external', href: 'https://example.com/a' },
			}),
		]);
		const edited = await editVsdx(added.bytes, [
			link({ row: 'Row_1', hyperlink: { address: '', subAddress: 'Page-1', description: 'Home' } }),
		]);
		expect((await first(edited.bytes)).hyperlinks).toEqual([
			expect.objectContaining({
				description: 'Home',
				target: { kind: 'internal', subAddress: 'Page-1' },
			}),
		]);
		const second = await editVsdx(edited.bytes, [link()]);
		expect((await first(second.bytes)).hyperlinks?.map((item) => item.name)).toEqual([
			'Row_1',
			'Row_2',
		]);
		const removed = await editVsdx(second.bytes, [
			link({ row: 'Row_1', hyperlink: null }),
			link({ row: 'Row_2', hyperlink: null }),
		]);
		expect((await first(removed.bytes)).hyperlinks).toEqual([]);
		expect(await page(removed.bytes)).not.toContain('Hyperlink');
	});
	it('keeps an unsafe address inert and rejects invalid or unsupported targets', async () => {
		const bytes = await drawing();
		const unsafe = await editVsdx(bytes, [
			link({ hyperlink: { address: 'javascript:alert(1)', subAddress: '', description: '' } }),
		]);
		expect((await first(unsafe.bytes)).hyperlinks?.[0]?.target.kind).toBe('unresolved');
		for (const bad of [
			link({ hyperlink: { address: '', subAddress: '', description: 'x' } }),
			link({ hyperlink: { address: 'a\nb', subAddress: '', description: '' } }),
			link({ hyperlink: null }),
			link({ row: 'bad row' }),
			link({ hyperlink: { address: 'x'.repeat(5000), subAddress: '', description: '' } }),
		])
			await expect(editVsdx(bytes, [bad])).rejects.toMatchObject({
				code: expect.stringMatching(/^INVALID_EDIT/),
			});
		await expect(editVsdx(bytes, [link({ row: 'Row_9' })])).rejects.toMatchObject({
			code: 'EDIT_TARGET_NOT_FOUND',
		});
		await expect(editVsdx(bytes, [link({ shapeId: '7' })])).rejects.toMatchObject({
			code: 'EDIT_TARGET_NOT_FOUND',
		});
	});
	it('refuses master instances, formula-driven rows and formulas that read the cells', async () => {
		const master = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${shape('1', rectangle, 'Master="2"')}</Shapes>` }],
			masters: [{ id: '2', shapes: shape('1') }],
		});
		await expect(editVsdx(master, [link()])).rejects.toMatchObject({
			code: 'UNSUPPORTED_METADATA_EDIT',
		});
		const formula = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', rectangle + '<Section N="Hyperlink"><Row N="Row_1">' + cell('Address', 'x', 'Prop.Url') + '</Row></Section>')}</Shapes>`,
				},
			],
		});
		await expect(editVsdx(formula, [link({ row: 'Row_1' })])).rejects.toMatchObject({
			code: 'UNSUPPORTED_METADATA_EDIT',
		});
		const reader = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1')}${shape('2', cell('Width', 1, 'LEN(Sheet.1!Comment)'))}</Shapes>`,
				},
			],
		});
		await expect(
			editVsdx(reader, [{ type: 'set-shape-screentip', pageId: '0', shapeId: '1', text: 'x' }]),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_METADATA_EDIT' });
	});
});

describe('shape ScreenTips', () => {
	it('sets, replaces and clears the Comment cell and reads it as the ScreenTip', async () => {
		const tip = (text: string) => ({
			type: 'set-shape-screentip' as const,
			pageId: '0',
			shapeId: '1',
			text,
		});
		const set = await editVsdx(await drawing(), [tip('Line one, "line" <two> & more')]);
		expect(await page(set.bytes)).toContain('N="Comment"');
		expect((await first(set.bytes)).screenTip).toBe('Line one, "line" <two> & more');
		await expect(editVsdx(set.bytes, [tip('two\nlines')])).rejects.toMatchObject({
			code: 'INVALID_EDIT_TEXT',
		});
		const replaced = await editVsdx(set.bytes, [tip('Short')]);
		expect((await first(replaced.bytes)).screenTip).toBe('Short');
		const unchanged = await editVsdx(replaced.bytes, [tip('Short')]);
		expect(unchanged.changedParts).toEqual([]);
		const cleared = await editVsdx(replaced.bytes, [tip('')]);
		expect((await first(cleared.bytes)).screenTip).toBeUndefined();
		expect(await page(cleared.bytes)).not.toContain('Comment');
		await expect(editVsdx(cleared.bytes, [tip('\u0001')])).rejects.toMatchObject({
			code: 'INVALID_EDIT_TEXT',
		});
	});
	it('works on an inserted picture so pictures can carry links and tips', async () => {
		const png = Uint8Array.from(
			Buffer.from(
				'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
				'base64',
			),
		);
		const picture = await editVsdx(await createVsdx(), [
			{
				type: 'insert-picture',
				pageId: '0',
				shapeId: '1',
				x: 1,
				y: 1,
				width: 1,
				height: 1,
				image: png,
			},
		]);
		const linked = await editVsdx(picture.bytes, [
			link(),
			{ type: 'set-shape-screentip', pageId: '0', shapeId: '1', text: 'Logo' },
		]);
		const result = await first(linked.bytes);
		expect(result).toMatchObject({ kind: 'foreign', screenTip: 'Logo' });
		expect(result.image?.mimeType).toBe('image/png');
		expect(result.hyperlinks).toHaveLength(1);
		const xml = await page(linked.bytes);
		expect(xml.indexOf('N="Hyperlink"')).toBeLessThan(xml.indexOf('<ForeignData'));
	});
});

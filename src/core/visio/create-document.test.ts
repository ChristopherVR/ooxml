import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { captureVisioClipboard } from './clipboard';
import { visioPasteCommand } from './ui/shape-clipboard';
import { VisioPackage } from './package';

describe('blank drawing creation', () => {
	it('creates an editable empty Letter page with consistent OPC references and declared fonts', async () => {
		const bytes = await createVsdx();
		const drawing = await parseVsdx(bytes);
		expect(drawing.format).toBe('vsdx');
		expect(drawing.pages).toHaveLength(1);
		expect(drawing.pages[0]).toMatchObject({
			id: '0',
			name: 'Page-1',
			width: 8.5,
			height: 11,
			shapes: [],
		});
		expect(drawing.fontFamilies).toContain('Calibri');
		const pkg = await VisioPackage.open(bytes);
		expect(pkg.paths()).toHaveLength(8);
		expect((await pkg.relationships('visio/document.xml')).get('rId2')).toMatchObject({
			target: 'visio/windows.xml',
		});
		expect((await editVsdx(bytes, [])).bytes).toEqual(bytes);
	});
	it('supports normal shape creation, formatting, source clipboard and reloading', async () => {
		const original = await createVsdx({ width: 11, height: 8.5 });
		const drawn = await editVsdx(original, [
			{
				type: 'create-rectangle',
				pageId: '0',
				shapeId: '1',
				x: 2,
				y: 3,
				width: 2,
				height: 1,
				text: 'New diagram',
			},
		]);
		const formatted = await editVsdx(drawn.bytes, [
			{ type: 'format-text', pageId: '0', shapeId: '1', bold: true },
		]);
		const drawing = await parseVsdx(formatted.bytes);
		expect(drawing.pages[0]!.shapes[0]!.text.plainText).toBe('New diagram');
		const capture = await captureVisioClipboard(formatted.bytes, '0', ['1']);
		const pasted = await editVsdx(formatted.bytes, [
			visioPasteCommand(drawing.pages[0]!, capture)!,
		]);
		expect(
			(await parseVsdx(pasted.bytes)).pages[0]!.shapes.map((shape) => shape.text.plainText),
		).toEqual(['New diagram', 'New diagram']);
		const cleared = await editVsdx(pasted.bytes, [
			{ type: 'delete-shape', pageId: '0', shapeId: '1' },
			{ type: 'delete-shape', pageId: '0', shapeId: '2' },
		]);
		expect((await parseVsdx(cleared.bytes)).pages[0]!.shapes).toHaveLength(0);
	});
	it.each([0, -1, NaN, Infinity, 1_000_001])(
		'rejects invalid physical dimensions: %s',
		async (width) => {
			await expect(createVsdx({ width })).rejects.toMatchObject({ code: 'INVALID_PAGE_SIZE' });
			await expect(createVsdx({ height: width })).rejects.toMatchObject({
				code: 'INVALID_PAGE_SIZE',
			});
		},
	);
	it('owns dimensions before asynchronous package writing', async () => {
		const options = { width: 10, height: 6 };
		const pending = createVsdx(options);
		options.width = 1;
		options.height = 1;
		expect((await parseVsdx(await pending)).pages[0]).toMatchObject({ width: 10, height: 6 });
	});
});

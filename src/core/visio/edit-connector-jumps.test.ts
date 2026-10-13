import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, shape } from './test-fixtures';
import { copySnapshotScene } from './ui/snapshot-scene';

const target = { type: 'format-shape' as const, pageId: '0', shapeId: '1' };
const source = (extra = '') =>
	fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', cell('OneD', 1) + extra)}</Shapes>` }],
	});
describe('connector jump formatting', () => {
	it('saves all Shape Layout settings, preserves other parts and copies them into print scenes', async () => {
		const bytes = await source();
		const patch = {
			lineJumpCode: 2,
			lineJumpStyle: 9,
			lineJumpDirectionX: 2,
			lineJumpDirectionY: 1,
		};
		const saved = await editVsdx(bytes, [{ ...target, ...patch }]);
		const scene = await parseVsdx(saved.bytes);
		expect(scene.pages[0]!.shapes[0]).toMatchObject(patch);
		expect(copySnapshotScene(scene).pages[0]!.shapes[0]).toMatchObject(patch);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
		expect((await editVsdx(saved.bytes, [{ ...target, ...patch }])).changedParts).toEqual([]);
		const restored = await editVsdx(saved.bytes, [
			{
				...target,
				lineJumpCode: 0,
				lineJumpStyle: 0,
				lineJumpDirectionX: 0,
				lineJumpDirectionY: 0,
			},
		]);
		expect((await parseVsdx(restored.bytes)).pages[0]!.shapes[0]!.lineJumpCode).toBe(0);
		const original = await JSZip.loadAsync(bytes),
			next = await JSZip.loadAsync(saved.bytes);
		for (const name of Object.keys(original.files).filter(
			(name) => !original.files[name]!.dir && name !== 'visio/pages/page1.xml',
		))
			expect(await next.file(name)!.async('uint8array')).toEqual(
				await original.file(name)!.async('uint8array'),
			);
	});
	it.each([
		{ lineJumpCode: 5 },
		{ lineJumpStyle: 10 },
		{ lineJumpDirectionX: -1 },
		{ lineJumpDirectionY: 1.5 },
	])('rejects invalid enums %j', async (patch) => {
		await expect(editVsdx(await source(), [{ ...target, ...patch }])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
	});
	it('refuses 2D shapes and guarded settings', async () => {
		await expect(editVsdx(await fixture(), [{ ...target, lineJumpCode: 2 }])).rejects.toMatchObject(
			{ code: 'INVALID_EDIT' },
		);
		await expect(
			editVsdx(await source(cell('ConLineJumpCode', 1, 'GUARD(1)')), [
				{ ...target, lineJumpCode: 2 },
			]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
	});
	it('preserves native self-transform triggers but refuses foreign or nested trigger dependencies', async () => {
		const trigger = cell('BegTrigger', 1, '_XFTRIGGER(Sheet.1!EventXFMod)');
		const saved = await editVsdx(await source(trigger), [{ ...target, lineJumpCode: 2 }]);
		const xml = await (
			await JSZip.loadAsync(saved.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		expect(xml).toContain('_XFTRIGGER(Sheet.1!EventXFMod)');
		for (const extra of [
			cell('BegTrigger', 1, '_XFTRIGGER(Sheet.2!EventXFMod)'),
			`<Section N="User"><Row N="x">${trigger}</Row></Section>`,
		]) {
			await expect(
				editVsdx(await source(extra), [{ ...target, lineJumpCode: 2 }]),
			).rejects.toMatchObject({
				code: expect.stringMatching(/^EDIT_(UNSUPPORTED_FORMAT|UNKNOWN)_DEPENDENCY$/),
			});
		}
	});
	it('overrides a master policy without changing the master or the other instance', async () => {
		const bytes = await fixture({
			masters: [{ id: '2', shapes: shape('1', cell('OneD', 1) + cell('ConLineJumpCode', 1)) }],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', '', 'Master="2"')}${shape('2', '', 'Master="2"')}</Shapes>`,
				},
			],
		});
		const saved = await editVsdx(bytes, [{ ...target, lineJumpCode: 3 }]);
		expect(
			(await parseVsdx(saved.bytes)).pages[0]!.shapes.map((shape) => shape.lineJumpCode),
		).toEqual([3, 1]);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
	});
	it('accepts native self triggers with namespace-prefixed page roots', async () => {
		const zip = await JSZip.loadAsync(
			await source(cell('BegTrigger', 1, '_XFTRIGGER(Sheet.1!EventXFMod)')),
		);
		const xml = await zip.file('visio/pages/page1.xml')!.async('string');
		zip.file(
			'visio/pages/page1.xml',
			xml
				.replace(
					'<PageContents ',
					'<v:PageContents xmlns:v="http://schemas.microsoft.com/office/visio/2012/main" ',
				)
				.replace('</PageContents>', '</v:PageContents>'),
		);
		const saved = await editVsdx(await zip.generateAsync({ type: 'uint8array' }), [
			{ ...target, lineJumpCode: 2 },
		]);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.lineJumpCode).toBe(2);
	});
});

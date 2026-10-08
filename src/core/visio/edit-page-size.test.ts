import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { createVsdx } from './create-document';
import { fixture, cell, shape } from './test-fixtures';
import { snapshotEdits } from './ui/edit-commands';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import {
	visioPageSizeCommand,
	visioPageSizeState,
	visioPageOrientationCommand,
	visioPageSizePresetCommand,
} from './ui/page-size';
import type { VisioPageSizeEdit } from './edit-commands';

const command: VisioPageSizeEdit = {
	type: 'set-page-size',
	pageId: '0',
	width: 6.25,
	height: 4.75,
};
const fixed = cell('DrawingSizeType', 3) + cell('DrawingResizeType', 0);
const scales = (scale = 1) => cell('PageScale', 1) + cell('DrawingScale', scale);
const contents = `<Shapes>${shape('1', cell('PinX', 20) + cell('PinY', -2) + cell('Width', 3) + cell('Height', 1) + cell('Angle', 0.5) + '<Text>Unchanged text</Text>')}</Shapes>`;
const source = (pageCells = fixed + scales(), shapeContents = contents, document = '') =>
	fixture({
		pages: [{ id: '0', contents: shapeContents, pageCells }],
		document,
		edit: (zip) => zip.file('custom/opaque.bin', new Uint8Array([0, 255, 1])),
	});
async function mutate(bytes: Uint8Array, path: string, change: (xml: string) => string) {
	const zip = await JSZip.loadAsync(bytes);
	zip.file(path, change(await zip.file(path)!.async('string')));
	return zip.generateAsync({ type: 'uint8array' });
}

describe('fixed drawing-page size', () => {
	it.each([0.5, 1, 2])(
		'writes physical inches at drawing scale %s and preserves every other payload',
		async (scale) => {
			const bytes = await source(fixed + scales(scale));
			const before = await parseVsdx(bytes),
				saved = await editVsdx(bytes, [command]);
			const after = await parseVsdx(saved.bytes),
				first = after.pages[0]!;
			expect(first).toMatchObject({
				width: 6.25,
				height: 4.75,
				drawingSizeType: 3,
				drawingResizeType: 0,
			});
			expect(first.shapes).toEqual(before.pages[0]!.shapes);
			expect(saved.changedParts).toEqual(['visio/pages/pages.xml']);
			const oldZip = await JSZip.loadAsync(bytes),
				newZip = await JSZip.loadAsync(saved.bytes);
			for (const path of Object.keys(oldZip.files))
				if (!oldZip.files[path]!.dir && path !== 'visio/pages/pages.xml')
					expect(await newZip.file(path)!.async('uint8array')).toEqual(
						await oldZip.file(path)!.async('uint8array'),
					);
			expect(await newZip.file('visio/pages/pages.xml')!.async('string')).toContain(
				`N="PageWidth" V="${6.25 * scale}"`,
			);
		},
	);
	it('converts a new drawing to fixed/custom without manufacturing a dirty no-op', async () => {
		const bytes = await createVsdx();
		const page = (await parseVsdx(bytes)).pages[0]!;
		expect(page.drawingSizeType).toBe(0);
		expect(page.drawingResizeType).toBeUndefined();
		expect(visioPageSizeCommand(page, page.width, page.height)).toHaveLength(1);
		const saved = await editVsdx(bytes, [{ ...command, width: page.width, height: page.height }]);
		expect(saved.changedParts).toEqual(['visio/pages/pages.xml']);
		const fixedPage = (await parseVsdx(saved.bytes)).pages[0]!;
		expect(visioPageSizeCommand(fixedPage, fixedPage.width, fixedPage.height)).toEqual([]);
		const noop = await editVsdx(saved.bytes, [
			{ ...command, width: fixedPage.width, height: fixedPage.height },
		]);
		expect(noop.bytes).toEqual(saved.bytes);
		expect(noop.changedParts).toEqual([]);
	});
	it('retains independent background page formulas, printer setup, units and unknown XML', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents,
					pageCells: fixed + scales(2) + cell('PrintPageOrientation', 2) + cell('PaperKind', 9),
					attributes: 'BackPage="1"',
				},
				{
					id: '1',
					contents: `<Shapes>${shape('1', cell('Width', 1, 'ThePage!PageWidth/5'))}</Shapes>`,
					attributes: 'Background="1"',
					pageCells: fixed + scales(2),
				},
			],
		});
		const modified = await mutate(bytes, 'visio/pages/pages.xml', (xml) =>
			xml
				.replace('N="PageWidth" V="8.5"', 'N="PageWidth" V="8.5" U="MM"')
				.replace('</PageSheet>', '<Unknown xmlns="urn:opaque">retain</Unknown></PageSheet>'),
		);
		const saved = await editVsdx(modified, [command]);
		const zip = await JSZip.loadAsync(saved.bytes),
			oldZip = await JSZip.loadAsync(modified);
		expect(await zip.file('visio/pages/page2.xml')!.async('uint8array')).toEqual(
			await oldZip.file('visio/pages/page2.xml')!.async('uint8array'),
		);
		const pages = await zip.file('visio/pages/pages.xml')!.async('string');
		expect(pages).toContain('N="PrintPageOrientation" V="2"');
		expect(pages).toContain('N="PaperKind" V="9"');
		expect(pages).toContain('N="PageWidth" V="12.5" U="MM"');
		expect(pages).toContain('xmlns="urn:opaque">retain');
	});
	it('compares exact physical no-op values before conversion noise', async () => {
		const bytes = await source(fixed + cell('PageScale', 0.7) + cell('DrawingScale', 1.3));
		const page = (await parseVsdx(bytes)).pages[0]!;
		const saved = await editVsdx(bytes, [{ ...command, width: page.width, height: page.height }]);
		expect(saved.changedParts).toEqual([]);
		expect(saved.bytes).toEqual(bytes);
	});
	it.each(['ThePage!PageWidth/2', 'PageHeight/2', 'INDIRECT(&quot;ThePage!PageWidth&quot;)'])(
		'refuses affected or dynamic local dependency %s',
		async (formula) => {
			await expect(
				editVsdx(
					await source(undefined, `<Shapes>${shape('1', cell('PinX', 1, formula))}</Shapes>`),
					[command],
				),
			).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY' });
		},
	);
	it.each(['PageWidth', 'DrawingSizeType', 'DrawingResizeType'])(
		'refuses inherited %s cache dependencies',
		async (name) => {
			await expect(
				editVsdx(
					await source(
						cell('DrawingSizeType', 0) + cell('DrawingResizeType', 1) + scales(),
						undefined,
						`<StyleSheets><StyleSheet ID="0">${cell('Width', 1, `ThePage!${name}`)}</StyleSheet></StyleSheets>`,
					),
					[command],
				),
			).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY' });
		},
	);
	it('refuses static cross-page references and unknown formula-bearing Visio metadata', async () => {
		const bytes = await fixture({
			pages: [
				{ id: '0', contents, pageCells: fixed + scales() },
				{
					id: '1',
					contents: `<Shapes>${shape('1', cell('Width', 8.5, 'Pages[Page 1]!ThePage!PageWidth'))}</Shapes>`,
				},
			],
		});
		await expect(editVsdx(bytes, [command])).rejects.toMatchObject({
			code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY',
		});
		const zip = await JSZip.loadAsync(await source());
		zip.file('visio/custom.xml', '<Extra><Value F="ThePage!PageWidth"/></Extra>');
		await expect(
			editVsdx(await zip.generateAsync({ type: 'uint8array' }), [command]),
		).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY' });
	});
	it.each(['PageWidth', 'PageScale', 'DrawingSizeType', 'DrawingResizeType'])(
		'refuses guarded %s without touching source',
		async (name) => {
			const bytes = await mutate(await source(), 'visio/pages/pages.xml', (xml) =>
				xml.replace(new RegExp(`N="${name}" V="([^"]+)"`), `N="${name}" V="$1" F="GUARD($1)"`),
			);
			await expect(editVsdx(bytes, [command])).rejects.toMatchObject({
				code: 'EDIT_PROTECTED_CELL',
			});
		},
	);
	it.each(['F="Inh"', 'F="1"', 'E="1"', 'U="STR"'])(
		'refuses unresolved or stale page caches %s',
		async (attributes) => {
			const bytes = await mutate(await source(), 'visio/pages/pages.xml', (xml) =>
				xml.replace('N="PageWidth" V="8.5"', `N="PageWidth" V="8.5" ${attributes}`),
			);
			await expect(editVsdx(bytes, [command])).rejects.toThrow();
		},
	);
	it.each([0, -1, NaN, Infinity, 1_000_001])(
		'refuses invalid physical dimensions %s',
		async (width) => {
			await expect(editVsdx(await source(), [{ ...command, width }])).rejects.toMatchObject({
				code: 'INVALID_EDIT',
			});
		},
	);
	it('owns commands before await and bounds converted raw drawing dimensions', async () => {
		const mutable = { ...command },
			bytes = await source();
		const pending = editVsdx(bytes, [mutable]);
		mutable.width = 100;
		expect((await parseVsdx((await pending).bytes)).pages[0]!.width).toBe(6.25);
		await expect(editVsdx(await source(fixed + scales(1e6)), [command])).rejects.toMatchObject({
			code: 'INVALID_PAGE_SIZE',
		});
		expect(snapshotEdits([command])).toEqual([command]);
	});
	it('preserves atomicity when a later page command is refused', async () => {
		const bytes = await source();
		await expect(
			editVsdx(bytes, [command, { ...command, pageId: 'absent' }]),
		).rejects.toMatchObject({ code: 'EDIT_TARGET_NOT_FOUND' });
		expect((await parseVsdx(bytes)).pages[0]!.width).toBe(8.5);
	});
	it('owns physical preset/orientation policy and copies validated mode caches', async () => {
		const page = (await parseVsdx(await source())).pages[0]!;
		expect(visioPageSizeState(page)?.orientation).toBe('portrait');
		expect(visioPageOrientationCommand(page, 'landscape')).toEqual([
			{ ...command, width: 11, height: 8.5 },
		]);
		expect(visioPageSizePresetCommand({ ...page, width: 11, height: 8.5 }, 'a4')).toEqual([
			{ ...command, width: 297 / 25.4, height: 210 / 25.4 },
		]);
		expect(visioPageOrientationCommand({ ...page, width: 5, height: 5 }, 'landscape')).toEqual([]);
		expect(visioPageSizeCommand(page, -1, 1)).toBeUndefined();
		expect(visioPageSizePresetCommand(page, 'unknown')).toBeUndefined();
		const scene = await parseVsdx(await source());
		const copy = copySnapshotScene(scene);
		expect(copy.pages[0]).toMatchObject({ drawingSizeType: 3, drawingResizeType: 0 });
		assertViewableDocument(copy);
		copy.pages[0]!.drawingSizeType = NaN;
		expect(() => assertViewableDocument(copy)).toThrow('mode');
	});
	it.each(['V="invalid"', 'V="3" U="STR"', 'V="3" E="1"', 'V="3" F="Inh"', 'V="256"'])(
		'omits an invalid explicit mode cache %s without guessing a no-op',
		async (attributes) => {
			const bytes = await mutate(await source(), 'visio/pages/pages.xml', (xml) =>
				xml.replace('N="DrawingSizeType" V="3"', `N="DrawingSizeType" ${attributes}`),
			);
			const page = (await parseVsdx(bytes)).pages[0]!;
			expect(page.drawingSizeType).toBeUndefined();
			expect(visioPageSizeCommand(page, page.width, page.height)).toHaveLength(1);
		},
	);
	it('refuses duplicated and case-variant page cells', async () => {
		const bytes = await source();
		for (const transform of [
			(xml: string) => xml.replace('</PageSheet>', cell('PageWidth', 8.5) + '</PageSheet>'),
			(xml: string) => xml.replace('N="PageWidth"', 'N="pagewidth"'),
		])
			await expect(
				editVsdx(await mutate(bytes, 'visio/pages/pages.xml', transform), [command]),
			).rejects.toMatchObject({ code: 'EDIT_AMBIGUOUS_CELL' });
	});
	it('refuses mixed shape/page edits before committing either operation', async () => {
		await expect(
			editVsdx(await source(), [command, { type: 'delete-shape', pageId: '0', shapeId: '1' }]),
		).rejects.toMatchObject({ code: 'EDIT_MIXED_PAGE_TRANSACTION' });
	});
	it.each(['PARENT()', 'DEPENDSON(&quot;ThePage!PageWidth&quot;)', 'UNKNOWNFUNCTION()'])(
		'refuses hidden dependencies without explicit page-cell tokens: %s',
		async (formula) => {
			await expect(
				editVsdx(
					await source(undefined, `<Shapes>${shape('1', cell('PinX', 1, formula))}</Shapes>`),
					[command],
				),
			).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY' });
		},
	);
});

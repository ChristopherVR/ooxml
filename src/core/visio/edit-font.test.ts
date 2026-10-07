import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { editVsdx } from './edit.js';
import { parseVsdx } from './parser.js';
import { cell, fixture, shape, rectangle } from './test-fixtures.js';

const fontFixture = (formula: string) =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape(
					'1',
					cell('Width', 4) +
						cell('Height', 2) +
						rectangle +
						`<Section N="Character"><Row IX="0">${cell('Font', 0, formula.replaceAll('"', '&quot;'))}</Row></Section>`,
				)}</Shapes>`,
			},
		],
	});
const resize = { type: 'resize-shape' as const, pageId: '0', shapeId: '1', width: 8, height: 4 };

it('preserves independent native font lookup formulas during geometry editing', async () => {
	const saved = await editVsdx(await fontFixture('FONT("Arial")'), [resize]);
	const zip = await JSZip.loadAsync(saved.bytes);
	expect(await zip.file('visio/pages/page1.xml')!.async('string')).toContain(
		'FONT(&quot;Arial&quot;)',
	);
	expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.width).toBe(8);
});
it('rejects affected font lookups instead of retaining stale font caches', async () => {
	await expect(
		editVsdx(await fontFixture('FONT(IF(Width&gt;1,"Arial","Calibri"))'), [resize]),
	).rejects.toThrow(/unsupported functions/i);
});

const native = process.env.VISIO_NATIVE_PAGE_SCALES_DIR;
describe.skipIf(!native)('native font lookup geometry edits', () => {
	it('edits a native scaled page without dropping its font formula', async () => {
		const source = await readFile(resolve(native!, 'page-scales.vsdx'));
		const saved = await editVsdx(source, [
			{ ...resize, pageId: '6', shapeId: '2' },
			{ type: 'move-shape', pageId: '6', shapeId: '2', x: 6, y: 4 },
		]);
		const model = await parseVsdx(saved.bytes);
		const page = model.pages.find((p) => p.id === '6')!;
		const rectangle = page.shapes.find((s) => s.id === '2')!;
		expect([rectangle.width, rectangle.height]).toEqual([4, 2]);
		expect(rectangle.transform.slice(4)).toEqual([1, 1]);
		expect(page.drawingToPageScale).toBe(0.5);
		const zip = await JSZip.loadAsync(saved.bytes);
		expect(await zip.file(saved.changedParts[0]!)!.async('string')).toContain(
			'FONT(&quot;Arial&quot;)',
		);
	});
});

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { editVsdx, parseVsdx } from './index';

const MiB = 1024 * 1024;

/** A page of `count` labelled rectangles in a row, each with its own outline. */
async function drawing(count: number): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await createVsdx({ width: 34, height: 22 }));
	const shapes = Array.from(
		{ length: count },
		(_, i) =>
			`<Shape ID="${i + 1}" Type="Shape"><Cell N="PinX" V="${(i % 100) * 0.3 + 0.2}"/><Cell N="PinY" V="${Math.floor(i / 100) * 0.3 + 0.2}"/><Cell N="Width" V="0.2"/><Cell N="Height" V="0.2"/><Cell N="LocPinX" V="0.1" F="Width*0.5"/><Cell N="LocPinY" V="0.1" F="Height*0.5"/><Section N="Geometry" IX="0"><Row T="MoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row><Row T="LineTo" IX="2"><Cell N="X" V="0.2"/><Cell N="Y" V="0"/></Row><Row T="LineTo" IX="3"><Cell N="X" V="0.2"/><Cell N="Y" V="0.2"/></Row><Row T="LineTo" IX="4"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row></Section><Text>S${i + 1}</Text></Shape>`,
	);
	zip.file(
		'visio/pages/page1.xml',
		`<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" xml:space="preserve"><Shapes>${shapes.join('')}</Shapes></PageContents>`,
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('memory budget of a read or an edit', () => {
	// 300 shapes cost about 9 MiB of XML and 1 MiB of model: a read fits 12 MiB, an edit does not.
	const limits = { maxXmlHeapBytes: 12 * MiB };

	it('opens a drawing within the budget and says when an edit of it would not fit', async () => {
		const document = await parseVsdx(await drawing(300), { limits });
		expect(document.pages[0]!.shapes).toHaveLength(300);
		expect(document.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'edit-size-limit',
				severity: 'info',
				message:
					'This drawing is too large to edit here: an edit would need more than 12 MB of memory.',
			}),
		);
		const roomy = await parseVsdx(await drawing(300), { limits: { maxXmlHeapBytes: 64 * MiB } });
		expect(roomy.diagnostics.some((item) => item.code === 'edit-size-limit')).toBe(false);
	});

	it('charges every package an edit opens to one budget', async () => {
		// Each package alone fits; the edit's copy of what it reads and its re-read of the result do not.
		const bytes = await drawing(300);
		await expect(
			editVsdx(bytes, [{ type: 'rename-page', pageId: '0', name: 'Renamed' }], { limits }),
		).rejects.toMatchObject({
			code: 'LIMIT_XML_MEMORY',
			message:
				'This drawing would need more than 12 MB of memory to edit, more than this viewer allows.',
		});
		const edited = await editVsdx(bytes, [{ type: 'rename-page', pageId: '0', name: 'Renamed' }], {
			limits: { maxXmlHeapBytes: 64 * MiB },
		});
		expect((await parseVsdx(edited.bytes)).pages[0]!.name).toBe('Renamed');
	});

	it('charges the model a read builds, not only its XML', async () => {
		// The page's XML costs about 1.68 MiB and the model of its 60 shapes 0.2 MiB more.
		const bytes = await drawing(60);
		await expect(
			parseVsdx(bytes, { limits: { maxXmlHeapBytes: 1.8 * MiB } }),
		).rejects.toMatchObject({
			code: 'LIMIT_XML_MEMORY',
			message:
				'This drawing would need more than 2 MB of memory to open, more than this viewer allows.',
		});
		await expect(parseVsdx(bytes, { limits: { maxXmlHeapBytes: 2 * MiB } })).resolves.toBeDefined();
	});

	it('says a drawing whose pages hold more cells than an edit follows is too large to edit', async () => {
		const document = await parseVsdx(await drawing(4_600), {
			limits: { maxRuntimeMs: 120_000, maxXmlHeapBytes: 4096 * MiB },
		});
		expect(document.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'edit-size-limit',
				message:
					'This drawing is too large to edit here: its pages hold about 101200 ShapeSheet cells, more than the 100000 an edit can follow.',
			}),
		);
	}, 60_000);
});

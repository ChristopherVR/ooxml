import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { parseVsdx } from './index';

/** A page of `count` plain rectangles with a label, the shape the UI audit measured. */
async function largePage(count: number): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await createVsdx({ width: 34, height: 22 }));
	const cols = Math.ceil(Math.sqrt(count * 1.5));
	const width = 34 / cols;
	const height = 22 / Math.ceil(count / cols);
	const shapes: string[] = [];
	for (let i = 0; i < count; i++) {
		const w = (width * 0.7).toFixed(4);
		const h = (height * 0.6).toFixed(4);
		const x = ((i % cols) * width + width / 2).toFixed(4);
		const y = (22 - (Math.floor(i / cols) * height + height / 2)).toFixed(4);
		shapes.push(
			`<Shape ID="${i + 1}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0"><Cell N="PinX" V="${x}"/><Cell N="PinY" V="${y}"/><Cell N="Width" V="${w}"/><Cell N="Height" V="${h}"/><Cell N="LocPinX" V="${Number(w) / 2}" F="Width*0.5"/><Cell N="LocPinY" V="${Number(h) / 2}" F="Height*0.5"/><Cell N="LineWeight" V="0.01"/><Section N="Geometry" IX="0"><Row T="MoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row><Row T="LineTo" IX="2"><Cell N="X" V="${w}"/><Cell N="Y" V="0"/></Row><Row T="LineTo" IX="3"><Cell N="X" V="${w}"/><Cell N="Y" V="${h}"/></Row><Row T="LineTo" IX="4"><Cell N="X" V="0"/><Cell N="Y" V="${h}"/></Row><Row T="LineTo" IX="5"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row></Section><Text>S${i + 1}</Text></Shape>`,
		);
	}
	zip.file(
		'visio/pages/page1.xml',
		`<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" xml:space="preserve"><Shapes>${shapes.join('')}</Shapes></PageContents>`,
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('large pages', () => {
	it(
		'opens a page of 10,000 shapes within the default XML limits',
		{ timeout: 120_000 },
		async () => {
			// The time budget is the host's (measured in docs); a loaded test runner must not trip it.
			const document = await parseVsdx(await largePage(10_000), {
				limits: { maxRuntimeMs: 120_000 },
			});
			expect(document.pages[0]!.shapes).toHaveLength(10_000);
			expect(document.pages[0]!.shapes.at(-1)!.text.plainText).toBe('S10000');
		},
	);
	it('refuses a larger page with its shape count', { timeout: 120_000 }, async () => {
		await expect(
			parseVsdx(await largePage(18_000), { limits: { maxRuntimeMs: 120_000 } }),
		).rejects.toMatchObject({
			code: 'LIMIT_XML',
			message: expect.stringMatching(/^The page has 18000 shapes, more than this viewer can open/),
		});
	});
});

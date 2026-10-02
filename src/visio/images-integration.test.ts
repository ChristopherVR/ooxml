import { describe, expect, it } from 'vitest';
import { parseVsdx } from './index.js';
import { cell, fixture, relations, shape } from './test-fixtures.js';
const pixel = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
		'base64',
	),
);
const foreign =
	'<ForeignData ForeignType="Bitmap" CompressionType="PNG"><Rel r:id="image1"/></ForeignData>';
const imageRel =
	'<Relationship Id="image1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.png"/>';
describe('embedded VSDX raster integration', () => {
	it('loads a validated embedded raster with its cached placement', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', cell('Width', 4) + cell('Height', 2) + cell('ImgOffsetX', 0.2) + cell('ImgOffsetY', 0.3) + cell('ImgWidth', 3) + cell('ImgHeight', 1) + foreign, 'Type="Foreign"')}</Shapes>`,
				},
			],
			edit: (zip) => {
				zip.file('visio/pages/_rels/page1.xml.rels', relations(imageRel));
				zip.file('visio/media/image.png', pixel);
			},
		});
		const document = await parseVsdx(bytes),
			item = document.pages[0]!.shapes[0]!;
		expect(item.image).toMatchObject({
			mimeType: 'image/png',
			pixelWidth: 1,
			pixelHeight: 1,
			x: 0.2,
			y: 0.3,
			width: 3,
			height: 1,
		});
		expect(item.image!.bytes).toEqual(pixel);
		expect(document.diagnostics.some((d) => d.code === 'unsupported-foreign-object')).toBe(false);
	});
	it('resolves master-owned image relationships before instance expansion and shares payloads', async () => {
		const bytes = await fixture({
			masters: [{ id: '5', shapes: shape('10', foreign) }],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', '', 'Master="5"')}${shape('2', '', 'Master="5"')}</Shapes>`,
				},
			],
			edit: (zip) => {
				zip.file('visio/masters/_rels/master1.xml.rels', relations(imageRel));
				zip.file('visio/media/image.png', pixel);
			},
		});
		const document = await parseVsdx(bytes),
			[first, second] = document.pages[0]!.shapes;
		expect(first!.image!.bytes).toBe(second!.image!.bytes);
	});
	it('omits unsupported raster payloads and respects image byte limits', async () => {
		const bytes = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${shape('1', foreign)}</Shapes>` }],
			edit: (zip) => {
				zip.file('visio/pages/_rels/page1.xml.rels', relations(imageRel));
				zip.file('visio/media/image.png', pixel);
			},
		});
		const document = await parseVsdx(bytes, { images: { maxImageBytes: 8 } });
		expect(document.pages[0]!.shapes[0]!.image).toBeUndefined();
		expect(document.diagnostics.some((d) => d.code === 'image-limit')).toBe(true);
	});
});

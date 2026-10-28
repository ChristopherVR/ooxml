import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx } from './index';
import { createDocument } from './model';
import { expectParagraph } from './test-support/access';
import { shapeSvg } from './shape-svg';
import { custom } from '../drawingml/test-support/custom-geometry';
it('renders, resizes and edits a custom Word shape without flattening its guides or opaque children', async () => {
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p',
			runs: [
				{
					text: '',
					image: {
						relId: '',
						partName: '',
						contentType: 'application/octet-stream',
						widthPx: 200,
						heightPx: 80,
						unsupported: 'Shape',
						shape: { preset: 'custom', customGeometryXml: custom, editable: true },
					},
				},
			],
		},
	];
	const loaded = await loadDocx(await saveDocx(model));
	const image = expectParagraph(loaded.model.blocks[0]).runs[0]!.image!;
	expect(image.shape).toMatchObject({ preset: 'custom', editable: true });
	expect(shapeSvg(image.shape!, 200, 80)).toContain('M50 0 L200 80');
	image.widthPx = 400;
	image.shape!.fill = { kind: 'solid', color: { kind: 'srgb', value: '123456', transforms: [] } };
	const bytes = await loaded.save(loaded.model);
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	expect(xml).toContain('fmla="*/ w adj 100000"');
	expect(xml).toContain('value="keep"');
	expect(xml).toContain('val="123456"');
	const reopened = await loadDocx(bytes);
	const again = expectParagraph(reopened.model.blocks[0]).runs[0]!.image!;
	expect(again.widthPx).toBe(400);
	expect(shapeSvg(again.shape!, 400, 80)).toContain('M100 0 L400 80');
});

import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { GLB_PART_CONTENT_TYPE, MODEL3D_REL } from '../drawingml/model3d';
import { embeddableOfficePackageFormat } from '../opc/embeddable-office-package';
import { RELATIONSHIP_TYPES, listEmbeddedPackages } from '../opc/index';
import { createDocument, loadDocx, saveDocx } from './index';
import { expectParagraph, must } from './test-support/access';

it('saves native GLB and Office payloads with preview images, then keeps both after text edits', async () => {
	const document = createDocument();
	const paragraph = expectParagraph(document.blocks[0]);
	const png = Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10);
	const modelPart = 'word/media/model.glb';
	const objectPart = 'word/embeddings/nested.docx';
	const child = await saveDocx(createDocument());
	const glb = Uint8Array.of(0x67, 0x6c, 0x54, 0x46, 0x02);
	const format = must(embeddableOfficePackageFormat('docx'));
	const preview = {
		relId: '',
		contentType: 'image/png',
		widthPx: 96,
		heightPx: 96,
	};
	paragraph.runs = [
		{ text: 'Before' },
		{
			text: '',
			image: {
				...preview,
				partName: 'word/media/model.png',
				title: 'Model.glb',
				payload: { kind: 'model3d', partName: modelPart },
			},
		},
		{
			text: '',
			image: {
				...preview,
				partName: 'word/media/object.png',
				title: 'Nested.docx',
				payload: { kind: 'officePackage', partName: objectPart, progId: format.progId },
			},
		},
		{ text: 'After' },
	];
	const pending = new Map([
		['word/media/model.png', { bytes: png, contentType: 'image/png' }],
		['word/media/object.png', { bytes: png, contentType: 'image/png' }],
		[modelPart, { bytes: glb, contentType: GLB_PART_CONTENT_TYPE }],
		[objectPart, { bytes: child, contentType: format.mime }],
	]);
	const inserted = await saveDocx(document, pending);
	const reopened = await loadDocx(inserted);
	const runs = expectParagraph(reopened.model.blocks[0]).runs;
	const model = must(runs[1]?.image);
	const object = must(runs[2]?.image);
	expect(model.payload).toEqual({ kind: 'model3d', partName: modelPart });
	expect(object.payload).toEqual({
		kind: 'officePackage',
		partName: objectPart,
		progId: format.progId,
	});
	expect(reopened.media?.get(model.partName)).toEqual(png);
	expect(reopened.media?.get(object.partName)).toEqual(png);

	runs[0] = { text: 'Updated' };
	model.widthPx = 120;
	object.widthPx = 80;
	const edited = await reopened.save();
	const zip = await JSZip.loadAsync(edited);
	const xml = await zip.file('word/document.xml')?.async('string');
	const rels = await zip.file('word/_rels/document.xml.rels')?.async('string');
	expect(xml).toContain('Requires="am3d"');
	expect(xml).toContain('<mc:Fallback>');
	expect(xml?.match(/<wp:extent cx="1143000"/g)).toHaveLength(2);
	expect(xml).toContain('DrawAspect="Icon"');
	expect(xml).toContain('width:60pt');
	expect(rels).toContain(MODEL3D_REL);
	expect(rels).toContain(RELATIONSHIP_TYPES.package);
	expect(await zip.file(modelPart)?.async('uint8array')).toEqual(glb);
	expect(await listEmbeddedPackages(edited)).toEqual([
		{ path: objectPart, name: 'nested.docx', kind: 'docx', bytes: child },
	]);
	const finalRuns = expectParagraph((await loadDocx(edited)).model.blocks[0]).runs;
	expect(finalRuns.map((run) => run.text).join('')).toBe('UpdatedAfter');
	expect(finalRuns[1]?.image?.widthPx).toBe(120);
	expect(finalRuns[2]?.image?.widthPx).toBe(80);
});

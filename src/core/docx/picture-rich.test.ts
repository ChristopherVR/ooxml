import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { NS, parseXml, buildXml, first } from '../xml/index';
import { applyPictureLayout } from '../drawingml/picture-layout';
import { applyPictureCrop } from '../drawingml/picture-crop';
import { applyPictureEffects, patchPictureEffects } from '../drawingml/picture-effects';
it('Word reads and saves shared tiled picture formatting without losing original media or opaque XML', async () => {
	const zip = await JSZip.loadAsync(
		readFileSync(
			new URL('./__fixtures__/review-inline/picture-insert-accepted.docx', import.meta.url),
		),
	);
	const document = parseXml(await zip.file('word/document.xml')!.async('string'));
	const fill = document.getElementsByTagNameNS(NS.pic, 'blipFill')[0]!;
	const blip = first(fill, 'blip', NS.a)!;
	fill.setAttributeNS('urn:vendor', 'v:opaque', 'keep');
	applyPictureLayout(fill, {
		tile: { scaleX: 2, scaleY: 0.5, offsetXEmu: 9525, alignment: 'ctr', flip: 'xy' },
		rotWithShape: false,
	});
	applyPictureCrop(fill, { cropLeft: 0.1 });
	applyPictureEffects(blip, [
		{
			kind: 'standard',
			name: 'duotone',
			colors: [
				{ kind: 'srgb', value: '000000', transforms: [] },
				{ kind: 'srgb', value: 'FF0000', transforms: [] },
			],
		},
		{ kind: 'standard', name: 'alphaBiLevel', values: { threshold: 0.4 } },
		{ kind: 'unsupported', name: 'vendor', sourceXml: '<v:vendor xmlns:v="urn:vendor"/>' },
	]);
	zip.file('word/document.xml', buildXml(document));
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const image = loaded.model.blocks.flatMap((block) =>
		block.type === 'paragraph' ? block.runs.flatMap((run) => (run.image ? [run.image] : [])) : [],
	)[0]!;
	expect(image.pictureFill).toMatchObject({
		crop: { cropLeft: 0.1 },
		layout: { rotWithShape: false, tile: { scaleX: 2, flip: 'xy' } },
		effects: [
			{ kind: 'standard', name: 'duotone' },
			{ kind: 'standard', name: 'alphaBiLevel' },
			{ kind: 'unsupported' },
		],
	});
	image.pictureFill!.layout!.tile!.scaleX = 0.75;
	image.pictureFill!.effects = patchPictureEffects(image.pictureFill!.effects!, {
		standard: { name: 'alphaBiLevel', values: { threshold: 0.6 } },
	});
	const saved = await loaded.save(loaded.model);
	const reopened = await loadDocx(saved);
	const again = reopened.model.blocks.flatMap((block) =>
		block.type === 'paragraph' ? block.runs.flatMap((run) => (run.image ? [run.image] : [])) : [],
	)[0]!;
	expect(again.pictureFill?.layout?.tile?.scaleX).toBe(0.75);
	expect(again.pictureFill?.effects?.[1]).toMatchObject({
		kind: 'standard',
		name: 'alphaBiLevel',
		values: { threshold: 0.6 },
	});
	const output = await JSZip.loadAsync(saved);
	const xml = await output.file('word/document.xml')!.async('string');
	expect(xml).toContain('v:opaque="keep"');
	expect(xml).toContain('<v:vendor');
	expect(await output.file(image.partName)!.async('uint8array')).toEqual(
		await zip.file(image.partName)!.async('uint8array'),
	);
});

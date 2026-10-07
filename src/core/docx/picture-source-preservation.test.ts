import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { rejectAllRevisions } from './revision-commands';
import { expectParagraph } from './test-support/access';
import { parseXml, buildXml, WORD_NS } from './xml';

const pictureFixture = () =>
	readFile(
		new URL('./__fixtures__/review-object-formatting/picture-tracked.docx', import.meta.url),
	);
function drawings(xml: string): string[] {
	return Array.from(parseXml(xml).getElementsByTagNameNS(WORD_NS, 'drawing')).map(buildXml);
}
it('preserves native drawing and opaque properties after text splits before a rejected picture format change', async () => {
	const source = await pictureFixture();
	const loaded = await loadDocx(source);
	const resolved = rejectAllRevisions(loaded.model);
	const paragraph = expectParagraph(resolved.blocks[0]);
	const first = paragraph.runs[0]!;
	paragraph.runs.splice(0, 1, { ...first, text: 'Be' }, { ...first, text: 'fore', italic: true });
	const before = await JSZip.loadAsync(source);
	const bytes = await loaded.save(resolved);
	const after = await JSZip.loadAsync(bytes);
	const xml = await after.file('word/document.xml')!.async('string');
	expect(drawings(xml)).toEqual(drawings(await before.file('word/document.xml')!.async('string')));
	expect(xml).toContain('noProof');
	expect(xml).not.toContain('rPrChange');
	expect(await after.file('word/media/image1.png')!.async('uint8array')).toEqual(
		await before.file('word/media/image1.png')!.async('uint8array'),
	);
	const reopened = await loadDocx(bytes);
	const runs = expectParagraph(reopened.model.blocks[0]).runs;
	expect(runs.map((run) => run.text).join('')).toBe('BeforeAfter');
	expect(runs.filter((run) => run.image)).toHaveLength(1);
	expect(runs.find((run) => run.text === 'fore')?.italic).toBe(true);
});

it('preserves a native picture when the preceding text run is removed', async () => {
	const source = await pictureFixture();
	const loaded = await loadDocx(source);
	const resolved = rejectAllRevisions(loaded.model);
	expectParagraph(resolved.blocks[0]).runs.splice(0, 1);
	const before = await JSZip.loadAsync(source);
	const bytes = await loaded.save(resolved);
	const after = await JSZip.loadAsync(bytes);
	expect(drawings(await after.file('word/document.xml')!.async('string'))).toEqual(
		drawings(await before.file('word/document.xml')!.async('string')),
	);
	expect(
		expectParagraph((await loadDocx(bytes)).model.blocks[0])
			.runs.map((run) => run.text)
			.join(''),
	).toBe('After');
});

async function repeatedPictures() {
	const zip = await JSZip.loadAsync(await pictureFixture());
	const xml = parseXml(await zip.file('word/document.xml')!.async('string'));
	const drawing = xml.getElementsByTagNameNS(WORD_NS, 'drawing')[0]!;
	const first = drawing.parentNode!;
	const second = first.cloneNode(true) as Element;
	const properties = second.getElementsByTagNameNS(
		'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
		'docPr',
	)[0]!;
	properties.setAttribute('id', '2');
	properties.setAttribute('name', 'Second source drawing');
	first.parentNode!.insertBefore(second, first.nextSibling);
	zip.file('word/document.xml', buildXml(xml));
	return zip.generateAsync({ type: 'uint8array' });
}

it('retains both source drawings when text splits before repeated media', async () => {
	const source = await repeatedPictures();
	const loaded = await loadDocx(source);
	const resolved = rejectAllRevisions(loaded.model);
	const paragraph = expectParagraph(resolved.blocks[0]);
	const first = paragraph.runs[0]!;
	paragraph.runs.splice(0, 1, { ...first, text: 'Be' }, { ...first, text: 'fore', italic: true });
	const after = await JSZip.loadAsync(await loaded.save(resolved));
	const before = await JSZip.loadAsync(source);
	expect(drawings(await after.file('word/document.xml')!.async('string'))).toEqual(
		drawings(await before.file('word/document.xml')!.async('string')),
	);
});

it('guards ambiguous repeated-media removal and leaves the source package unchanged', async () => {
	const source = await repeatedPictures();
	const loaded = await loadDocx(source);
	const resolved = rejectAllRevisions(loaded.model);
	const paragraph = expectParagraph(resolved.blocks[0]);
	paragraph.runs.splice(
		paragraph.runs.findIndex((run) => run.image),
		1,
	);
	await expect(loaded.save(resolved)).rejects.toThrow('pictures sharing the same media');
	expect(await loaded.save()).toEqual(source);
});

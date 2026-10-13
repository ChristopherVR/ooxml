import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { PptxHandler } from '../../index';

describe('paragraph spacing inheritance (#74)', () => {
	it('keeps direct spacing on one paragraph through an edit, save and reload', async () => {
		const zip = await JSZip.loadAsync(
			readFileSync(fileURLToPath(new URL('../fixtures/e2e/linked-textbox.pptx', import.meta.url))),
		);
		const source = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const body = `<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:pPr><a:lnSpc><a:spcPts val="1286"/></a:lnSpc><a:spcAft><a:spcPts val="380"/></a:spcAft></a:pPr><a:r><a:t>First</a:t></a:r></a:p>
<a:p><a:r><a:t>Second</a:t></a:r></a:p><a:p><a:r><a:t>Third</a:t></a:r></a:p></p:txBody>`;
		zip.file('ppt/slides/slide1.xml', source.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/g, body));
		const handler = new PptxHandler();
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		const slide = data.slides[0]!;
		const shape = slide.elements.find(
			(element) => 'text' in element && element.text?.includes('First'),
		)!;
		expect(
			'textStyle' in shape &&
				shape.textStyle?.inheritedParagraphSpacing?.[0]?.paragraphSpacingAfter,
		).toBe(0);
		if ('textSegments' in shape) shape.textSegments![0]!.text = 'Edited first';
		if ('text' in shape) shape.text = shape.text!.replace('First', 'Edited first');
		slide.isDirty = true;
		const saved = await handler.save(data.slides);
		const savedXml = await (
			await JSZip.loadAsync(saved)
		)
			.file('ppt/slides/slide1.xml')!
			.async('string');
		const firstBody = savedXml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/)![0];
		expect((firstBody.match(/<a:lnSpc>/g) ?? []).length).toBe(1);
		expect((firstBody.match(/<a:spcAft>/g) ?? []).length).toBe(1);
		const reloaded = await new PptxHandler().load(saved);
		const edited = reloaded.slides[0]!.elements.find(
			(element) => 'text' in element && element.text?.includes('Edited first'),
		)!;
		expect('textStyle' in edited && edited.textStyle?.inheritedParagraphSpacing?.[0]).toEqual({
			lineSpacing: 1,
			paragraphSpacingBefore: 0,
			paragraphSpacingAfter: 0,
		});
	});
});

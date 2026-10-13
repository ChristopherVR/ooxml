import JSZip from 'jszip';
import { PresentationBuilder } from 'pptx-viewer-core';

/** Native PresentationML shape-paint target, leaving its text outside the background build. */
export async function backgroundAnimationDeck(): Promise<Buffer> {
	const { handler, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addShape('rect', {
			x: 150,
			y: 140,
			width: 500,
			height: 220,
			fill: { type: 'solid', color: '#00AA00' },
		})
		.build();
	const element = slide.elements[0]!;
	if (element.type !== 'shape') throw new Error('Expected shape');
	element.text = 'Persistent text';
	element.textSegments = [{ text: element.text, style: { fontSize: 36, color: '#000000' } }];
	slide.animations = [
		{ elementId: element.id, entrance: 'fadeIn', durationMs: 1000, trigger: 'onClick' },
	];
	const zip = await JSZip.loadAsync(await handler.save([slide]));
	handler.dispose();
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const patched = xml.replace(
		/<p:spTgt([^>]*?)(?:\/>|>\s*<\/p:spTgt>)/gu,
		'<p:spTgt$1><p:bg/></p:spTgt>',
	);
	if (patched === xml) throw new Error('No native animation shape target');
	zip.file('ppt/slides/slide1.xml', patched);
	return zip.generateAsync({ type: 'nodebuffer' });
}

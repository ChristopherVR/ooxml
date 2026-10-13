import JSZip from 'jszip';
import { NS, buildXml, elements, parseXml } from '../../../xml';
import { PresentationBuilder } from 'ooxml-core/pptx';

const PIXEL =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRz0AAAAASUVORK5CYII=';
/** Two decks with colliding native part names, distinct themes and source-only assets. */
export async function createSlideClipboardDeck(source: boolean): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create({
		theme: { colors: { accent1: source ? '#ABCDEF' : '#112233' } },
	});
	const first = createSlide('Blank').addText(source ? 'Source copied' : 'Destination first', {
		x: 50,
		y: 50,
		width: 300,
		height: 60,
		fontSize: 24,
	});
	if (source) first.addImage(PIXEL, { x: 100, y: 160, width: 100, height: 100 });
	data.slides = [
		first.build(),
		createSlide('Blank')
			.addText(source ? 'Private unselected slide' : 'Destination last')
			.build(),
	];
	if (source) {
		data.slides[0]!.notes = 'Source rich speaker notes';
		data.slides[0]!.notesSegments = [{ text: 'Source rich speaker notes', style: { bold: true } }];
	}
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const path = 'ppt/slides/slide1.xml';
	const doc = parseXml(await zip.file(path)!.async('string'));
	for (const properties of Array.from(doc.getElementsByTagNameNS(NS.a, 'rPr'))) {
		for (const child of elements(properties))
			if (child.localName === 'solidFill') properties.removeChild(child);
		const fill = doc.createElementNS(NS.a, 'a:solidFill'),
			scheme = doc.createElementNS(NS.a, 'a:schemeClr');
		scheme.setAttribute('val', 'accent1');
		fill.appendChild(scheme);
		properties.insertBefore(fill, properties.firstChild);
	}
	zip.file(path, buildXml(doc));
	return zip.generateAsync({ type: 'arraybuffer' });
}

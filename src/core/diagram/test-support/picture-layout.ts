import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { NS, parseXml, buildXml, first, children, type XmlElement } from '../../xml/index';
import { RELATIONSHIP_TYPES, buildRelationshipsXml } from '../../opc/index';

/** Native Office gallery layout/presentation identities, populated with actual image data. */
export async function withNativePictureLayout(
	bytes: Uint8Array,
	prefix: 'word' | 'xl' | 'ppt',
	layout = 'picture-accent-list--flat3',
): Promise<Uint8Array> {
	const native = await JSZip.loadAsync(
		readFileSync(
			new URL(`../../pptx/__tests__/fixtures/smartart-gallery/${layout}.pptx`, import.meta.url),
		),
	);
	const zip = await JSZip.loadAsync(bytes);
	const types = parseXml(await zip.file('[Content_Types].xml')!.async('string'));
	const typesNs = 'http://schemas.openxmlformats.org/package/2006/content-types';
	if (![...types.documentElement.childNodes].some(node => node.nodeType === 1 && (node as XmlElement).getAttribute('Extension') === 'png')) {
		const png = types.createElementNS(typesNs, 'Default');
		png.setAttribute('Extension', 'png');
		png.setAttribute('ContentType', 'image/png');
		types.documentElement.appendChild(png);
		zip.file('[Content_Types].xml', buildXml(types));
	}
	const data = parseXml(await native.file('ppt/diagrams/data1.xml')!.async('string'));
	const drawing = parseXml(await native.file('ppt/diagrams/drawing1.xml')!.async('string'));
	const points = children(first(data.documentElement, 'ptLst', NS.dgm)!, 'pt', NS.dgm);
	const pictures = points.filter(
		(point) =>
			first(point, 'prSet', NS.dgm)?.getAttribute('presStyleLbl')?.includes('Img') ||
			first(point, 'prSet', NS.dgm)?.getAttribute('presName') === 'image',
	);
	const rels = new Map();
	let count = 0;
	for (const point of pictures) {
		const id = point.getAttribute('modelId');
		const assoc = first(point, 'prSet', NS.dgm)?.getAttribute('presAssocID');
		const shape = [...drawing.getElementsByTagNameNS(NS.dsp, 'sp')].find(
			(shape) => shape.getAttribute('modelId') === id,
		);
		if (!shape || !assoc) continue;
		const rel = `picture${++count}`;
		const fill = parseXml(
			`<a:blipFill xmlns:a="${NS.a}" xmlns:r="${NS.r}"><a:blip r:embed="${rel}"><a:grayscl/><a:alphaModFix amt="75000"/></a:blip><a:srcRect l="12500" r="25000"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
		).documentElement;
		for (const owner of [shape, points.find((entry) => entry.getAttribute('modelId') === assoc)!]) {
			const props = first(owner, 'spPr', owner.namespaceURI!)!;
			for (const old of [...props.childNodes]
				.filter((child): child is XmlElement => child.nodeType === 1)
				.filter((child) =>
					['solidFill', 'noFill', 'blipFill', 'gradFill'].includes(child.localName),
				))
				props.removeChild(old);
			props.appendChild(props.ownerDocument.importNode(fill, true));
		}
		rels.set(rel, {
			type: RELATIONSHIP_TYPES.image,
			target: '../media/shared-reflow-picture.png',
			mode: 'Internal' as const,
		});
	}
	if (!count) throw new Error(`Native layout has no admitted picture role: ${layout}`);
	zip.file(`${prefix}/diagrams/data1.xml`, buildXml(data));
	zip.file(`${prefix}/diagrams/drawing1.xml`, buildXml(drawing));
	for (const name of ['layout', 'colors', 'quickStyle'])
		zip.file(
			`${prefix}/diagrams/${name}1.xml`,
			await native.file(`ppt/diagrams/${name}1.xml`)!.async('string'),
		);
	for (const name of ['data', 'drawing'])
		zip.file(`${prefix}/diagrams/_rels/${name}1.xml.rels`, buildRelationshipsXml(rels));
	zip.file(
		`${prefix}/media/shared-reflow-picture.png`,
		'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP8zwACTGCSAQANHQEDgslx/wAAAABJRU5ErkJggg==',
		{ base64: true },
	);
	return zip.generateAsync({ type: 'uint8array' });
}

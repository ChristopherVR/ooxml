import { NS, first, relAttr, type XmlElement } from '../xml/index';

const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const SVG_EXTENSION = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';

/** The Office SVG relationship supplementing the picture's raster blip. */
export function svgPictureRelationship(blip: XmlElement | undefined): string | undefined {
	const svg = blip?.getElementsByTagNameNS(SVG_NS, 'svgBlip')[0];
	return relAttr(svg, 'embed') || undefined;
}

export function setSvgPictureRelationship(blip: XmlElement, id: string): void {
	const doc = blip.ownerDocument;
	if (!doc) return;
	let extensions = first(blip, 'extLst', NS.a);
	if (!extensions) {
		extensions = doc.createElementNS(NS.a, 'a:extLst');
		blip.appendChild(extensions);
	}
	const existing = extensions.getElementsByTagNameNS(SVG_NS, 'svgBlip')[0];
	if (existing) {
		existing.setAttributeNS(NS.r, 'r:embed', id);
		return;
	}
	const extension = doc.createElementNS(NS.a, 'a:ext');
	extension.setAttribute('uri', SVG_EXTENSION);
	const svg = doc.createElementNS(SVG_NS, 'asvg:svgBlip');
	svg.setAttributeNS(NS.r, 'r:embed', id);
	extension.appendChild(svg);
	extensions.appendChild(extension);
}

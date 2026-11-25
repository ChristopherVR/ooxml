import { inspectSvgPicture, svgPictureBytes } from '../drawingml/static-svg-picture';
import { RASTER_IMAGE_LIMITS } from '../drawingml/raster-image';
import { findOfficeDocumentPart } from '../opc/package';
import { declaredEntrySize, openZip } from '../opc/zip-limits';
import { NS, buildXml, first, parseXml } from '../xml/index';
import type { VisioImage } from './model';
import type { VisioPackage } from './package';
import { children, type Report } from './sheet';

const MAX_EMBED_BYTES = 2 * 1024 * 1024;
const MAX_EMBEDDINGS = 16;
const MAX_EQUATIONS = 16;
const SVG = 'http://www.w3.org/2000/svg';
const previews = new WeakMap<VisioPackage, Map<string, Promise<VisioImage | undefined>>>();

/** Visio stores Word equations in an internal OLE package. Read math without activating OLE. */
export async function readEmbeddedEquationPreview(
	pkg: VisioPackage,
	source: string,
	foreignData: Element,
	report: Report,
): Promise<VisioImage | undefined> {
	if (foreignData.getAttribute('ShowAsIcon') === '1') return;
	const refs = children(foreignData, 'Rel');
	const id = refs.length === 1 ? refs[0]?.getAttributeNS(NS.r, 'id') : undefined;
	const rel = id ? (await pkg.relationships(source)).get(id) : undefined;
	if (rel?.mode !== 'Internal' || rel.type !== `${NS.r}/package` || !/\.docx$/i.test(rel.target))
		return;
	let cache = previews.get(pkg);
	if (!cache) previews.set(pkg, (cache = new Map()));
	let pending = cache.get(rel.target);
	if (!pending) {
		if (cache.size >= MAX_EMBEDDINGS || pkg.getPartByteLength(rel.target) > MAX_EMBED_BYTES) {
			report(
				'embedded-equation-limit',
				'An embedded equation document exceeded the preview limits.',
			);
			return;
		}
		pending = pkg.readBytes(rel.target).then(drawEquations);
		cache.set(rel.target, pending);
	}
	return pending;
}

async function drawEquations(bytes: Uint8Array): Promise<VisioImage | undefined> {
	try {
		const zip = await openZip(bytes, {
			maxInputBytes: MAX_EMBED_BYTES,
			maxEntries: 256,
			maxEntryBytes: 4 * MAX_EMBED_BYTES,
			maxTotalBytes: 8 * MAX_EMBED_BYTES,
		});
		const partName = await findOfficeDocumentPart(zip);
		const part = partName ? zip.file(partName.replace(/^\//, '')) : null;
		if (!part || declaredEntrySize(part) > MAX_EMBED_BYTES) return;
		const word = parseXml(await part.async('string'), {
			label: 'Embedded equation document',
			limits: { maxNodes: 30_000 },
		}).documentElement;
		if (word.localName !== 'document' || word.namespaceURI !== NS.w) return;
		const equations = Array.from(word.getElementsByTagNameNS(NS.m, 'oMath')).slice(
			0,
			MAX_EQUATIONS,
		);
		if (!equations.length) return;
		const { ommlFromElement, drawOmmlTree } = await import('ooxml-core/math');
		const doc = parseXml(`<svg xmlns="${SVG}"/>`).documentElement;
		let width = 0;
		let height = 0;
		for (const equation of equations) {
			try {
				const drawing = drawOmmlTree(ommlFromElement(equation), 20);
				if (!drawing || drawing.svg.length > 128 * 1024) continue;
				const svg = parseXml(drawing.svg).documentElement;
				const equationGroup = first(svg, 'g', SVG);
				if (!equationGroup) continue;
				for (const text of Array.from(equationGroup.getElementsByTagNameNS(SVG, 'text'))) {
					text.removeAttribute('textLength');
					text.removeAttribute('lengthAdjust');
				}
				const line = doc.ownerDocument.createElementNS(SVG, 'g');
				line.setAttribute('transform', `translate(0 ${height})`);
				line.appendChild(doc.ownerDocument.importNode(equationGroup, true));
				doc.appendChild(line);
				width = Math.max(width, drawing.widthPx);
				height += drawing.heightPx + 4;
			} catch {
				// Over-limit equations keep their source bytes and leave the preview available to others.
			}
		}
		if (!width || !height) return;
		doc.setAttribute('viewBox', `0 0 ${width} ${height}`);
		doc.setAttribute('preserveAspectRatio', 'xMinYMin meet');
		const preview = new TextEncoder().encode(buildXml(doc));
		const { root, pixelWidth, pixelHeight } = inspectSvgPicture(preview, RASTER_IMAGE_LIMITS);
		return { bytes: svgPictureBytes(root), mimeType: 'image/svg+xml', pixelWidth, pixelHeight };
	} catch {
		return undefined;
	}
}

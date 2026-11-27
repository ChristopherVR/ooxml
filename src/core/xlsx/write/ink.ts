import { INK_RELATIONSHIP_TYPE, INKML_CONTENT_TYPE } from '../../drawingml/inkml/constants';
import { NS, buildXml, parseXml } from '../../xml/index';
import type { InkObject } from '../model';
import { XDR14_NS } from '../read/ink';
import type { PackageWriter, RelationshipSet } from './package-writer';
import { escapeAttr } from './xml-out';

const EMU_PER_PIXEL = 9525;
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** A part name for new ink that neither the output nor the source package uses yet. */
function newInkPart(writer: PackageWriter): string {
	let index = 1;
	while (
		writer.has(`xl/ink/ink${index}.xml`) ||
		writer.source?.bytes(`xl/ink/ink${index}.xml`) !== undefined
	)
		index++;
	return `xl/ink/ink${index}.xml`;
}

/** Sets the size `xdr14:xfrm` gives a kept ink drawing, the space its strokes are written in. */
function patchXfrm(xml: string, box: { w: number; h: number }): string {
	let doc: ReturnType<typeof parseXml>;
	try {
		doc = parseXml(xml, { label: 'XLSX drawing' });
	} catch {
		return xml;
	}
	for (const xfrm of Array.from(doc.getElementsByTagNameNS(XDR14_NS, 'xfrm'))) {
		const ext = xfrm.getElementsByTagNameNS(NS.a, 'ext')[0];
		ext?.setAttribute('cx', String(Math.max(1, Math.round(box.w * EMU_PER_PIXEL))));
		ext?.setAttribute('cy', String(Math.max(1, Math.round(box.h * EMU_PER_PIXEL))));
	}
	return buildXml(doc);
}

/**
 * The anchor content of new ink as Excel reads it: `xdr:contentPart` with `xdr14:` children, in an
 * `mc:Choice Requires="xdr14"` so a reader without Excel 2010 drawings skips it. No fallback
 * picture is written; Excel 16 opens it as ink.
 */
function contentPartXml(
	relId: string,
	id: number,
	name: string,
	box: { x: number; y: number; w: number; h: number },
): string {
	const emu = (px: number, min = 0) => Math.max(min, Math.round(px * EMU_PER_PIXEL));
	return (
		`<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice xmlns:xdr14="${XDR14_NS}" Requires="xdr14">` +
		`<xdr:contentPart xmlns:r="${NS.r}" r:id="${relId}"><xdr14:nvContentPartPr>` +
		`<xdr14:cNvPr id="${id}" name="${escapeAttr(name)}"/><xdr14:cNvContentPartPr/></xdr14:nvContentPartPr>` +
		`<xdr14:nvPr macro=""/><xdr14:xfrm><a:off x="${emu(box.x)}" y="${emu(box.y)}"/>` +
		`<a:ext cx="${emu(box.w, 1)}" cy="${emu(box.h, 1)}"/></xdr14:xfrm></xdr:contentPart>` +
		`</mc:Choice></mc:AlternateContent>`
	);
}

/**
 * Writes an ink drawing's part and returns its anchor content (new ink) or its kept anchor XML
 * (`kept`). The strokes are written in their own `width` x `height` space, which `xdr14:xfrm`
 * names; Excel stretches the ink onto the anchor, so moving or resizing it rewrites nothing.
 * Untouched ink keeps its part's bytes; changed strokes rewrite it; a second drawing of the same
 * part (a copy) gets a part of its own.
 */
export function writeInkDrawing(
	writer: PackageWriter,
	ink: InkObject,
	rels: RelationshipSet,
	relative: (part: string) => string,
	position: { x: number; y: number },
	nextId: () => number,
	keepSource: boolean,
): { kept: string } | { content: string } {
	const inkml = writer.inkml;
	if (!inkml) throw new Error('The InkML writer was not loaded for this save.');
	const { updatedInkPart, writeInkPart } = inkml;
	const size = { w: ink.width, h: ink.height };
	const shared = ink.partName !== undefined && writer.inkParts.has(ink.partName);
	if (ink.partName) writer.inkParts.add(ink.partName);
	if (ink.partName && ink.sourceXml && keepSource && !shared) {
		const source = writer.source?.text(ink.partName);
		const xml = updatedInkPart(source, ink.strokes, { width: size.w, height: size.h });
		if (xml === undefined) {
			writer.carry(ink.partName);
			return { kept: ink.sourceXml };
		}
		writer.add(ink.partName, xml, INKML_CONTENT_TYPE);
		return { kept: patchXfrm(ink.sourceXml, size) };
	}
	const part = newInkPart(writer);
	// A copy of ink read only in part copies its part; anything else is written from its strokes.
	const bytes = ink.truncated && ink.partName ? writer.source?.bytes(ink.partName) : undefined;
	writer.add(
		part,
		bytes ?? writeInkPart(ink.strokes, undefined, { width: size.w, height: size.h }),
		INKML_CONTENT_TYPE,
	);
	const relId = rels.ensure(INK_RELATIONSHIP_TYPE, relative(part));
	const id = nextId();
	return {
		content: contentPartXml(relId, id, ink.name ?? `Ink ${id - 1}`, { ...position, ...size }),
	};
}

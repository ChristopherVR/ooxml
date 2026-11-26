// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reads a WordprocessingML group shape (`wpg:wgp`) into a composite `ShapeSpec` so the editor,
// Print Layout and PDF export render grouped drawings as vector graphics instead of placeholders.
import type { GroupShapeChild, ShapeSpec } from './shape-model';
import { A_NS, parseShapeElement, WPS_NS } from './shape-parse';
import { EMU_PER_PIXEL } from './units';
import { first, getW, isElement, WORD_NS, type XmlElement } from './xml';
import { parseHalfPoints } from './simple-types';

export const WPG_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

function childNS(
	parent: XmlElement | undefined,
	ns: string,
	local: string,
): XmlElement | undefined {
	if (!parent) return undefined;
	for (const node of Array.from(parent.childNodes))
		if (isElement(node) && node.namespaceURI === ns && node.localName === local) return node;
	return undefined;
}

const emuPx = (value: string | null) => {
	const parsed = value ? Number(value) : NaN;
	return Number.isFinite(parsed) ? parsed / EMU_PER_PIXEL : 0;
};

function childBox(spPr: XmlElement | undefined): {
	xPx: number;
	yPx: number;
	widthPx: number;
	heightPx: number;
} {
	const xfrm = childNS(spPr, A_NS, 'xfrm');
	const off = childNS(xfrm, A_NS, 'off');
	const ext = childNS(xfrm, A_NS, 'ext');
	return {
		xPx: emuPx(off?.getAttribute('x') ?? null),
		yPx: emuPx(off?.getAttribute('y') ?? null),
		widthPx: Math.max(0, emuPx(ext?.getAttribute('cx') ?? null)),
		heightPx: Math.max(0, emuPx(ext?.getAttribute('cy') ?? null)),
	};
}

function parseGroupNode(wgp: XmlElement, depth = 0): ShapeSpec | undefined {
	if (depth >= 32) return undefined;
	const grpSpPr = childNS(wgp, WPG_NS, 'grpSpPr');
	const xfrm = childNS(grpSpPr, A_NS, 'xfrm');
	const chOff = childNS(xfrm, A_NS, 'chOff');
	const chExt = childNS(xfrm, A_NS, 'chExt');
	const off = childNS(xfrm, A_NS, 'off');
	const ext = childNS(xfrm, A_NS, 'ext');
	const groupChildren: GroupShapeChild[] = [];
	let groupCanUngroup = true;
	for (const node of Array.from(wgp.childNodes)) {
		if (!isElement(node)) continue;
		if (node.namespaceURI === WPS_NS && node.localName === 'wsp') {
			const parsed = parseShapeElement(node, false);
			if (!parsed) {
				groupCanUngroup = false;
				continue;
			}
			groupCanUngroup &&= Boolean(
				parsed.editable && !parsed.text?.length && !childNS(node, WPS_NS, 'cNvPr'),
			);
			const runs = parsed.text?.some(Boolean)
				? Array.from(node.getElementsByTagNameNS(WORD_NS, 'r'))
				: [];
			const sizes = new Set(
				runs.map((run) => parseHalfPoints(getW(first(first(run, 'rPr'), 'sz'), 'val'))),
			);
			const [size] = sizes;
			const shape = sizes.size === 1 && size ? { ...parsed, groupTextSizePt: size / 2 } : parsed;
			const box = childBox(childNS(node, WPS_NS, 'spPr'));
			groupChildren.push({ ...box, shape });
		} else if (
			node.namespaceURI === WPG_NS &&
			(node.localName === 'grpSp' || node.localName === 'wgp')
		) {
			groupCanUngroup = false;
			const nested = parseGroupNode(node, depth + 1);
			if (!nested) continue;
			const box = childBox(childNS(node, WPG_NS, 'grpSpPr'));
			groupChildren.push({ ...box, shape: nested });
		} else if (
			node.namespaceURI !== WPG_NS ||
			!['cNvGrpSpPr', 'grpSpPr'].includes(node.localName)
		) {
			groupCanUngroup = false;
		}
	}
	if (!groupChildren.length) return undefined;
	const chW = emuPx(chExt?.getAttribute('cx') ?? null) || emuPx(ext?.getAttribute('cx') ?? null);
	const chH = emuPx(chExt?.getAttribute('cy') ?? null) || emuPx(ext?.getAttribute('cy') ?? null);
	const chX = emuPx(chOff?.getAttribute('x') ?? off?.getAttribute('x') ?? null);
	const chY = emuPx(chOff?.getAttribute('y') ?? off?.getAttribute('y') ?? null);
	const maxX = groupChildren.reduce((max, item) => Math.max(max, item.xPx + item.widthPx), chX);
	const maxY = groupChildren.reduce((max, item) => Math.max(max, item.yPx + item.heightPx), chY);
	const rot = Number(xfrm?.getAttribute('rot'));
	return {
		preset: 'rect',
		fill: { kind: 'none' },
		line: { fill: { kind: 'none' } },
		groupChildren,
		...(groupCanUngroup ? { groupCanUngroup: true } : {}),
		groupExtent: {
			xPx: chX,
			yPx: chY,
			widthPx: Math.max(1 / EMU_PER_PIXEL, chW || maxX - chX),
			heightPx: Math.max(1 / EMU_PER_PIXEL, chH || maxY - chY),
		},
		...(Number.isFinite(rot) && rot !== 0
			? { rotation: Math.round((rot / 60000) * 100) / 100 }
			: {}),
		...(xfrm?.getAttribute('flipH') === '1' ? { flipH: true } : {}),
		...(xfrm?.getAttribute('flipV') === '1' ? { flipV: true } : {}),
		editable: false,
		textEditable: false,
	};
}

/** Parses a `wpg:wgp` group inside `a:graphicData`, or undefined when it holds no drawable shapes. */
export function parseGroupShape(graphicData: XmlElement | undefined): ShapeSpec | undefined {
	const wgp = graphicData?.getElementsByTagNameNS(WPG_NS, 'wgp')[0] as XmlElement | undefined;
	return wgp ? parseGroupNode(wgp) : undefined;
}

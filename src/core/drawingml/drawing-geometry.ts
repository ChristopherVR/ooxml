// DrawingML transform and geometry readers over the DOM (preset evaluation itself lives in the
// `geometry` area).
import { parseCustomGeometry } from './custom-geometry';
import { NS, booleanAttribute, children, first, type XmlElement } from './dom';
import type { DrawingFrame, DrawingPath } from './types';

const int = (element: XmlElement | null | undefined, name: string): number =>
	Number.parseInt(element?.getAttribute(name) ?? '', 10);

/** A frame from an element holding `a:off` and `a:ext` (`a:xfrm`, `dsp:txXfrm`). */
export function parseFrame(container: XmlElement | undefined): DrawingFrame | undefined {
	const off = first(container, 'off', NS.a);
	const ext = first(container, 'ext', NS.a);
	if (!off || !ext) return undefined;
	return {
		x: int(off, 'x') || 0,
		y: int(off, 'y') || 0,
		width: int(ext, 'cx') || 0,
		height: int(ext, 'cy') || 0,
	};
}

export interface ParsedTransform {
	frame: DrawingFrame;
	/** Degrees. */
	rotation?: number;
	flipHorizontal?: boolean;
	flipVertical?: boolean;
}

/** Reads `a:xfrm`; `undefined` without `a:off` and `a:ext`. */
export function parseTransform(xfrm: XmlElement | undefined): ParsedTransform | undefined {
	const frame = parseFrame(xfrm);
	if (!frame || !xfrm) return undefined;
	const rot = int(xfrm, 'rot');
	return {
		frame,
		...(Number.isFinite(rot) && rot !== 0 ? { rotation: rot / 60000 } : {}),
		...(booleanAttribute(xfrm, 'flipH') ? { flipHorizontal: true } : {}),
		...(booleanAttribute(xfrm, 'flipV') ? { flipVertical: true } : {}),
	};
}

export interface ParsedGeometry {
	geometry: string;
	adjustments?: Record<string, number>;
	paths?: DrawingPath[];
	/** Set when a geometry is present but not fully modelled. */
	issue?: string;
}

/** Reads `a:prstGeom` (name and adjustments) or `a:custGeom` (literal paths). */
export function parseDrawingGeometry(spPr: XmlElement | undefined): ParsedGeometry {
	const preset = first(spPr, 'prstGeom', NS.a);
	const custom = first(spPr, 'custGeom', NS.a);
	if (custom) {
		const frame = parseFrame(first(spPr, 'xfrm', NS.a));
		const paths = parseCustomGeometry(custom, frame?.width ?? 21600, frame?.height ?? 21600);
		return paths
			? { geometry: 'custom', paths }
			: { geometry: 'custom', issue: 'CUSTOM_GEOMETRY_UNSUPPORTED' };
	}
	const geometry = preset?.getAttribute('prst') || 'rect';
	const adjustments: Record<string, number> = {};
	const avLst = first(preset, 'avLst', NS.a);
	if (avLst)
		for (const guide of children(avLst, 'gd', NS.a)) {
			const name = guide.getAttribute('name');
			const match = /^val\s+(-?\d+)/u.exec(guide.getAttribute('fmla') ?? '');
			if (name && match?.[1]) adjustments[name] = Number.parseInt(match[1], 10);
		}
	return { geometry, ...(Object.keys(adjustments).length > 0 ? { adjustments } : {}) };
}

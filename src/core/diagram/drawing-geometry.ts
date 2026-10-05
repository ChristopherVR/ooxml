// DrawingML transform and geometry readers over the DOM. Diagram-independent; destined for
// `drawingml` (preset evaluation itself lives in the `geometry` area).
import { NS, booleanAttribute, children, first, type XmlElement } from './dom.js';
import type { DiagramFrame, DiagramPath, DiagramPathCommand } from './types.js';

const int = (element: XmlElement | null | undefined, name: string): number =>
	Number.parseInt(element?.getAttribute(name) ?? '', 10);

/** A frame from an element holding `a:off` and `a:ext` (`a:xfrm`, `dsp:txXfrm`). */
export function parseFrame(container: XmlElement | undefined): DiagramFrame | undefined {
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
	frame: DiagramFrame;
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

function parsePathCommands(path: XmlElement): DiagramPathCommand[] | undefined {
	const commands: DiagramPathCommand[] = [];
	const point = (element: XmlElement) => {
		const x = int(element, 'x');
		const y = int(element, 'y');
		return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : undefined;
	};
	for (const element of Array.from(path.childNodes)) {
		if (element.nodeType !== 1) continue;
		const node = element as XmlElement;
		const points = children(node, 'pt', NS.a).map(point);
		if (points.some((entry) => entry === undefined)) return undefined;
		const [p0, p1, p2] = points as { x: number; y: number }[];
		switch (node.localName) {
			case 'moveTo':
			case 'lnTo':
				if (!p0) return undefined;
				commands.push({ op: node.localName === 'moveTo' ? 'M' : 'L', ...p0 });
				break;
			case 'cubicBezTo':
				if (!p0 || !p1 || !p2) return undefined;
				commands.push({ op: 'C', x1: p0.x, y1: p0.y, x2: p1.x, y2: p1.y, x: p2.x, y: p2.y });
				break;
			case 'quadBezTo':
				if (!p0 || !p1) return undefined;
				commands.push({ op: 'Q', x1: p0.x, y1: p0.y, x: p1.x, y: p1.y });
				break;
			case 'arcTo': {
				const wR = int(node, 'wR');
				const hR = int(node, 'hR');
				const startAngle = int(node, 'stAng');
				const swingAngle = int(node, 'swAng');
				if (![wR, hR, startAngle, swingAngle].every(Number.isFinite)) return undefined;
				commands.push({
					op: 'A',
					wR,
					hR,
					startAngle: startAngle / 60000,
					swingAngle: swingAngle / 60000,
				});
				break;
			}
			case 'close':
				commands.push({ op: 'Z' });
				break;
		}
	}
	return commands;
}

export interface ParsedGeometry {
	geometry: string;
	adjustments?: Record<string, number>;
	paths?: DiagramPath[];
	/** Set when a geometry is present but not fully modelled. */
	issue?: string;
}

/** Reads `a:prstGeom` (name and adjustments) or `a:custGeom` (literal paths). */
export function parseDrawingGeometry(spPr: XmlElement | undefined): ParsedGeometry {
	const preset = first(spPr, 'prstGeom', NS.a);
	const custom = first(spPr, 'custGeom', NS.a);
	if (custom) {
		const paths: DiagramPath[] = [];
		for (const path of children(first(custom, 'pathLst', NS.a) ?? custom, 'path', NS.a)) {
			const commands = parsePathCommands(path);
			// Guide-named coordinates need the guide evaluator; report rather than guess.
			if (!commands) return { geometry: 'custom', issue: 'CUSTOM_GEOMETRY_GUIDES_UNSUPPORTED' };
			const width = int(path, 'w');
			const height = int(path, 'h');
			paths.push({
				commands,
				...(Number.isFinite(width) ? { width } : {}),
				...(Number.isFinite(height) ? { height } : {}),
				...(path.getAttribute('fill') ? { fill: path.getAttribute('fill') as string } : {}),
				...(path.getAttribute('stroke') === '0' || path.getAttribute('stroke') === 'false'
					? { stroke: false }
					: {}),
			});
		}
		return { geometry: 'custom', paths };
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

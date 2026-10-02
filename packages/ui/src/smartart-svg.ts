import { normalizeHexColor } from 'ooxml-core/color';
import type {
	DiagramColor,
	DiagramDrawing,
	DiagramDrawingShape,
	DiagramFill,
	DiagramPath,
} from 'ooxml-core/diagram';

const SVG_NS = 'http://www.w3.org/2000/svg';
const EMU_PER_PX = 9525;

export interface SmartArtRenderReport {
	shapeCount: number;
	/** Preset geometries drawn as a plain rectangle because no outline is available here. */
	approximatedGeometries: string[];
	/** Shapes that carry 3D scene data: drawn flat, the 3D is kept only in the source part. */
	flattened3d: number;
	/** Fill kinds drawn as a neutral placeholder (gradient is approximated by its first stop). */
	approximatedFills: string[];
}

/** Theme colours the host resolved (`accent1` -> `#4472c4`); unresolved scheme colours are grey. */
export type SchemeColors = Readonly<Record<string, string>>;

export function colorToCss(
	color: DiagramColor | undefined,
	scheme: SchemeColors,
): string | undefined {
	if (!color) return undefined;
	switch (color.kind) {
		case 'srgb':
			return normalizeHexColor(color.value, '#808080');
		case 'scheme':
			return scheme[color.value] ? normalizeHexColor(scheme[color.value], '#808080') : '#9ca3af';
		case 'system':
			return color.fallback ? normalizeHexColor(color.fallback, '#808080') : undefined;
		case 'preset':
			return /^[a-z]+$/i.test(color.value) ? color.value.toLowerCase() : '#808080';
		default:
			return '#9ca3af';
	}
}

function fillToCss(
	fill: DiagramFill | undefined,
	scheme: SchemeColors,
	report: SmartArtRenderReport,
): string {
	if (!fill) return 'none';
	if (fill.kind === 'none') return 'none';
	if (fill.kind === 'solid') return colorToCss(fill.color, scheme) ?? 'none';
	if (!report.approximatedFills.includes(fill.kind)) report.approximatedFills.push(fill.kind);
	if (fill.kind === 'gradient') return colorToCss(fill.stops[0]?.color, scheme) ?? 'none';
	return '#d1d5db';
}

/** Path coordinates are in the path's own units; scale them to the frame (px). */
function pathData(path: DiagramPath, widthPx: number, heightPx: number): string {
	const sx = path.width ? widthPx / path.width : 1;
	const sy = path.height ? heightPx / path.height : 1;
	const x = (v: number) => +(v * sx).toFixed(2);
	const y = (v: number) => +(v * sy).toFixed(2);
	const out: string[] = [];
	for (const c of path.commands) {
		if (c.op === 'M' || c.op === 'L') out.push(`${c.op}${x(c.x)} ${y(c.y)}`);
		else if (c.op === 'C')
			out.push(`C${x(c.x1)} ${y(c.y1)} ${x(c.x2)} ${y(c.y2)} ${x(c.x)} ${y(c.y)}`);
		else if (c.op === 'Q') out.push(`Q${x(c.x1)} ${y(c.y1)} ${x(c.x)} ${y(c.y)}`);
		else if (c.op === 'Z') out.push('Z');
		// Arc commands (`A`) need the current point and angle maths; dropped, reported by the caller.
	}
	return out.join('');
}

const hasArc = (shape: DiagramDrawingShape): boolean =>
	Boolean(shape.paths?.some((p) => p.commands.some((c) => c.op === 'A')));

function shapeElement(
	doc: Document,
	shape: DiagramDrawingShape,
	report: SmartArtRenderReport,
): SVGElement {
	const { frame } = shape;
	const w = frame.width / EMU_PER_PX;
	const h = frame.height / EMU_PER_PX;
	if (shape.paths?.length) {
		const el = doc.createElementNS(SVG_NS, 'g');
		for (const path of shape.paths) {
			const p = doc.createElementNS(SVG_NS, 'path');
			p.setAttribute('d', pathData(path, w, h));
			el.append(p);
		}
		if (hasArc(shape) && !report.approximatedGeometries.includes('custom:arc')) {
			report.approximatedGeometries.push('custom:arc');
		}
		return el;
	}
	if (shape.geometry === 'ellipse') {
		const e = doc.createElementNS(SVG_NS, 'ellipse');
		e.setAttribute('rx', String(w / 2));
		e.setAttribute('ry', String(h / 2));
		e.setAttribute('cx', String(w / 2));
		e.setAttribute('cy', String(h / 2));
		return e;
	}
	const rect = doc.createElementNS(SVG_NS, 'rect');
	rect.setAttribute('width', String(w));
	rect.setAttribute('height', String(h));
	if (shape.geometry === 'roundRect') {
		const adj = (shape.adjustments?.adj ?? 16667) / 100000;
		rect.setAttribute('rx', String(Math.min(w, h) * adj));
	} else if (shape.geometry !== 'rect' && !report.approximatedGeometries.includes(shape.geometry)) {
		report.approximatedGeometries.push(shape.geometry);
		rect.setAttribute('data-approximated', shape.geometry);
	}
	return rect;
}

/** Render a core `DiagramDrawing` into an `<svg>`: pure DOM construction, no markup injection. */
export function renderDiagramDrawing(
	doc: Document,
	drawing: DiagramDrawing,
	scheme: SchemeColors = {},
): { svg: SVGSVGElement; report: SmartArtRenderReport } {
	const report: SmartArtRenderReport = {
		shapeCount: drawing.shapes.length,
		approximatedGeometries: [],
		flattened3d: 0,
		approximatedFills: [],
	};
	const svg = doc.createElementNS(SVG_NS, 'svg');
	let right = 0;
	let bottom = 0;
	for (const shape of drawing.shapes) {
		const { frame } = shape;
		right = Math.max(right, (frame.x + frame.width) / EMU_PER_PX);
		bottom = Math.max(bottom, (frame.y + frame.height) / EMU_PER_PX);
		const group = doc.createElementNS(SVG_NS, 'g');
		group.dataset.modelId = shape.modelId;
		const cx = frame.width / EMU_PER_PX / 2;
		const cy = frame.height / EMU_PER_PX / 2;
		const fx = shape.flipHorizontal ? -1 : 1;
		const fy = shape.flipVertical ? -1 : 1;
		const transform = [`translate(${frame.x / EMU_PER_PX} ${frame.y / EMU_PER_PX})`];
		if (shape.rotation) transform.push(`rotate(${shape.rotation} ${cx} ${cy})`);
		if (fx !== 1 || fy !== 1) {
			transform.push(`translate(${cx} ${cy}) scale(${fx} ${fy}) translate(${-cx} ${-cy})`);
		}
		group.setAttribute('transform', transform.join(' '));
		const body = shapeElement(doc, shape, report);
		body.setAttribute('fill', fillToCss(shape.fill, scheme, report));
		const stroke =
			shape.line?.fill?.kind === 'solid' ? colorToCss(shape.line.fill.color, scheme) : undefined;
		body.setAttribute('stroke', stroke ?? 'none');
		if (stroke)
			body.setAttribute(
				'stroke-width',
				String(Math.max(0.5, (shape.line?.widthEmu ?? 9525) / EMU_PER_PX)),
			);
		group.append(body);
		if (shape.has3d) {
			report.flattened3d += 1;
			group.setAttribute('data-flattened-3d', '');
		}
		if (shape.text?.text) group.append(textElement(doc, shape, scheme));
		svg.append(group);
	}
	svg.setAttribute('viewBox', `0 0 ${Math.max(right, 1)} ${Math.max(bottom, 1)}`);
	return { svg, report };
}

function textElement(doc: Document, shape: DiagramDrawingShape, scheme: SchemeColors): SVGElement {
	const frame = shape.textFrame ?? {
		x: shape.frame.x,
		y: shape.frame.y,
		width: shape.frame.width,
		height: shape.frame.height,
	};
	const ox = (frame.x - shape.frame.x) / EMU_PER_PX;
	const oy = (frame.y - shape.frame.y) / EMU_PER_PX;
	const w = frame.width / EMU_PER_PX;
	const h = frame.height / EMU_PER_PX;
	const text = doc.createElementNS(SVG_NS, 'text');
	const first = shape.text?.paragraphs[0]?.runs[0];
	const sizePx = ((first?.sizePt ?? 12) * 96) / 72;
	text.setAttribute('x', String(ox + w / 2));
	text.setAttribute('y', String(oy + h / 2));
	text.setAttribute('text-anchor', 'middle');
	text.setAttribute('dominant-baseline', 'central');
	text.setAttribute('font-size', String(sizePx));
	text.setAttribute('fill', colorToCss(first?.color, scheme) ?? 'currentColor');
	if (first?.bold) text.setAttribute('font-weight', '700');
	if (first?.italic) text.setAttribute('font-style', 'italic');
	if (first?.typeface) text.setAttribute('font-family', first.typeface);
	text.textContent =
		shape.text?.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join(' ') ?? '';
	return text;
}
